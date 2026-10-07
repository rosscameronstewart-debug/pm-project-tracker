import ast
import gc
import json
import sqlite3
import tempfile
import threading
import unittest
from pathlib import Path
from urllib.request import Request, urlopen
from urllib.error import HTTPError
from http.server import ThreadingHTTPServer
from unittest.mock import patch

import app
import access_control as access


class AccessTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.TemporaryDirectory()
        cls.root = Path(cls.tmp.name)
        cls.patches = [patch.object(app,'DB_PATH',cls.root/'test.sqlite3'),
                       patch.object(app,'DATA_DIR',cls.root),
                       patch.object(app,'UPLOAD_DIR',cls.root/'uploads'),
                       patch.object(app,'BRAND_DIR',cls.root/'brand'),
                       patch.object(app,'seed_bid_tracker_from_workbook')]
        for p in cls.patches:
            p.start()
        app.init_db()
        with app.db() as con:
            con.execute("INSERT INTO users(id,username,password_hash,role,created_at) VALUES (101,'field1@example.com','test','Field PO','2026-10-06'),(102,'field2@example.com','test','Field PO','2026-10-06')")
            con.execute("INSERT INTO purchase_orders (id,po_number,job_number,vendor,description,requested_by_user_id,attachment_file,pickup_file,invoice_file,created_at,updated_at) VALUES (10,'TEST-10','JOB1','Test','Test',101,'own.pdf','pickup.pdf','office.pdf','2026-10-06','2026-10-06')")
            con.execute("INSERT INTO purchase_orders (id,po_number,job_number,vendor,description,requested_by_user_id,attachment_file,created_at,updated_at) VALUES (11,'TEST-11','JOB2','Test','Test',102,'other.pdf','2026-10-06','2026-10-06')")
            con.execute("INSERT INTO financial_reports (report_date,report_type,source_file,uploaded_at) VALUES ('2026-10-06','pnl','financial.pdf','2026-10-06')")
        for name in ('own.pdf','pickup.pdf','office.pdf','other.pdf','financial.pdf','orphan.pdf'):
            (app.UPLOAD_DIR/name).write_bytes(b'%PDF-1.4 test')
        cls.actor = None
        cls.permissions = {}
        cls.user_patch = patch.object(app,'current_user', side_effect=lambda _:cls.actor)
        cls.role_patch = patch.object(app,'role_permissions', side_effect=lambda _:cls.permissions)
        cls.user_patch.start()
        cls.role_patch.start()
        cls.server = ThreadingHTTPServer(('127.0.0.1',0),app.Handler)
        cls.thread = threading.Thread(target=cls.server.serve_forever,daemon=True)
        cls.thread.start()
        cls.url = 'http://127.0.0.1:%s' % cls.server.server_port

    @classmethod
    def tearDownClass(cls):
        cls.server.shutdown()
        cls.server.server_close()
        cls.thread.join()
        cls.user_patch.stop()
        cls.role_patch.stop()
        for p in reversed(cls.patches):
            p.stop()
        gc.collect()
        cls.tmp.cleanup()

    def setUp(self):
        type(self).actor = {'id':101,'role':'User','username':'test@example.com'}
        type(self).permissions = {}

    def grant(self, key, edit=0):
        type(self).permissions[key] = {'can_view':1,'can_edit':edit}

    def request(self, path, method='GET', data=None):
        try:
            body = (json.dumps(data).encode() if data is not None else b'') if method != 'GET' else None
            with urlopen(Request(self.url+path, data=body, method=method, headers={'Content-Type':'application/json'}),timeout=10) as response:
                return response.status, response.read()
        except HTTPError as error:
            return error.code, error.read()

    def test_all_existing_routes_have_a_policy(self):
        tree = ast.parse(Path(app.__file__).read_text(encoding='utf-8'))
        handler = next(n for n in tree.body if isinstance(n,ast.ClassDef) and n.name=='Handler')
        for function in handler.body:
            if function.name not in ('do_GET','do_POST','do_PUT','do_DELETE'):
                continue
            method = function.name[3:]
            for node in ast.walk(function):
                if isinstance(node,ast.Call) and isinstance(node.func,ast.Attribute) and ast.unparse(node.func.value)=='parsed.path' and node.func.attr=='startswith':
                    for argument in node.args:
                        if isinstance(argument,ast.Constant) and isinstance(argument.value,str) and argument.value.startswith('/api/') and argument.value != '/api/':
                            with self.subTest(method=method,prefix=argument.value):
                                self.assertNotEqual(access.requirement(method,argument.value+'1'),())
                if isinstance(node,ast.Compare) and ast.unparse(node.left)=='parsed.path':
                    for value in node.comparators:
                        constants = [value] if isinstance(value,ast.Constant) else value.elts if isinstance(value,ast.Tuple) else []
                        for constant in constants:
                            if not isinstance(constant,ast.Constant) or not isinstance(constant.value,str):
                                continue
                            path = constant.value
                            if path in ('/api/login','/login','/manifest.json','/service-worker.js','/offline'):
                                continue
                            with self.subTest(method=method,path=path):
                                self.assertNotEqual(access.requirement(method,path),())

    def test_feature_reads_and_writes_are_denied_without_permission(self):
        for path in ('/api/projects','/api/summary','/api/customer-invoices','/api/bid-summary',
                     '/api/texas-financial-summary','/api/internal-rates','/api/imports',
                     '/api/vendor-invoice-lines','/api/nte-summary','/api/admin/database-backup'):
            with self.subTest(path=path):
                self.assertEqual(self.request(path)[0],403)
        for method,path in (('POST','/api/projects'),('PUT','/api/projects/1'),('DELETE','/api/subprojects/1'),
                            ('POST','/api/import-fieldwise'),('POST','/api/customer-invoices'),
                            ('POST','/api/admin/send-test-email')):
            with self.subTest(method=method,path=path):
                self.assertEqual(self.request(path,method)[0],403)
        self.assertEqual(self.request('/api/unknown-new-route')[0],403)

    def test_read_only_can_read_but_cannot_mutate(self):
        type(self).actor['role']='Read Only'
        self.grant('bids')
        self.assertEqual(self.request('/api/bids')[0],200)
        for method,path in (('POST','/api/bids'),('PUT','/api/bids/1'),('DELETE','/api/subprojects/1')):
            self.assertEqual(self.request(path,method)[0],403)

    def test_document_download_preview_and_page_use_same_owner(self):
        type(self).actor['role']='Field PO'
        self.grant('po_requests',1)
        self.assertEqual(self.request('/uploads/own.pdf')[0],200)
        self.assertEqual(self.request('/po/10')[0],200)
        self.assertEqual(self.request('/po/11')[0],404)
        for name in ('other.pdf','office.pdf','financial.pdf','orphan.pdf'):
            for prefix,suffix in (('/uploads/',''),('/pdf-viewer/',''),('/pdf-page/','/0.png')):
                with self.subTest(name=name,prefix=prefix):
                    self.assertEqual(self.request(prefix+name+suffix)[0],404)
        for prefix,suffix in (('/pdf-viewer/',''),('/pdf-page/','/0.png')):
            self.assertTrue(access.document_allowed(app,type(self).actor,access.document_name(prefix+'own.pdf'+suffix)))
        self.assertEqual(self.request('/api/purchase-orders/11/pickup','POST')[0],404)
        self.assertEqual(self.request('/api/purchase-orders/10/invoice','POST')[0],403)
        body=json.loads(self.request('/api/purchase-orders')[1])
        self.assertEqual([po['id'] for po in body],[10])
        self.assertNotIn('invoices',body[0])
        self.assertNotIn(b'office.pdf',self.request('/po/10')[1])

    def test_office_and_texas_document_permissions(self):
        self.grant('po_review')
        self.assertEqual(self.request('/uploads/office.pdf')[0],200)
        self.assertEqual(self.request('/uploads/other.pdf')[0],200)
        self.assertEqual(self.request('/uploads/financial.pdf')[0],404)
        type(self).actor['role']='TX/Read Only'
        type(self).permissions={}
        self.grant('texas_ops')
        self.assertEqual(self.request('/uploads/financial.pdf')[0],200)
        self.assertEqual(self.request('/uploads/office.pdf')[0],404)

    def test_signed_out_and_no_data_changes_on_rejected_requests(self):
        type(self).actor=None
        self.assertEqual(self.request('/api/bids')[0],401)
        self.assertEqual(self.request('/api/projects','POST')[0],401)
        self.assertEqual(self.request('/api/projects/1','PUT')[0],401)
        self.assertEqual(self.request('/api/subprojects/1','DELETE')[0],401)
        self.assertEqual(app.one('SELECT count(*) AS n FROM projects')['n'],0)

    def test_disabled_account_and_revoked_permission(self):
        self.grant('bids')
        self.assertEqual(self.request('/api/bids')[0],200)
        type(self).permissions={}
        self.assertEqual(self.request('/api/bids')[0],403)
        type(self).user_patch.stop()
        token=app.create_session(101)
        try:
            app.execute('UPDATE users SET active=0 WHERE id=101')
            with self.assertRaises(HTTPError) as denied:
                urlopen(Request(self.url+'/api/me',headers={'Cookie':'pm_session='+token}),timeout=10)
            self.assertEqual(denied.exception.code,401)
        finally:
            app.execute('UPDATE users SET active=1 WHERE id=101')
            app.execute('DELETE FROM user_sessions WHERE session_token=?',(token,))
            type(self).user_patch.start()

    def test_mixed_import_deletion_checks_payload_source(self):
        self.grant('fieldwise',1)
        self.assertEqual(self.request('/api/imports/delete','POST',{'source':'Vendor Invoice'})[0],403)
        self.grant('vendor_invoices',1)
        self.assertEqual(self.request('/api/imports/delete','POST',{'source':'Unknown'})[0],403)

    def test_alerts_and_imports_do_not_leak_other_features(self):
        self.grant('bids')
        class Handler:
            command='GET'
            path='/api/home-alerts'
        payload={'alerts':[{'tab':'bids','severity':'warn','count':1},{'tab':'texasOps','severity':'bad','count':1}], 'attention_count':2}
        filtered=access.filter_payload(app,Handler(),payload)
        self.assertEqual(filtered['attention_count'],1)
        self.assertEqual(len(filtered['alerts']),1)
        Handler.path='/api/imports'
        self.grant('fieldwise')
        self.assertEqual(access.filter_payload(app,Handler(),[{'source':'Field Wise'},{'source':'Vendor Invoice'}]),[{'source':'Field Wise'}])


if __name__=='__main__':
    unittest.main()
