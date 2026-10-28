import gc
import json
import io
import sqlite3
import tempfile
import threading
import unittest
from contextlib import closing
from pathlib import Path
from http.server import ThreadingHTTPServer
from urllib.error import HTTPError
from urllib.request import Request, urlopen
from unittest.mock import patch

import app
import material_review as review
import material_reconciliation as rec


class IntegratedReviewTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name)
        self.patches = [patch.object(app, 'DATA_DIR', self.root),
                        patch.object(app, 'DB_PATH', self.root / 'app.sqlite3'),
                        patch.object(app, 'UPLOAD_DIR', self.root / 'uploads'),
                        patch.object(app, 'BRAND_DIR', self.root / 'brand'),
                        patch.object(app, 'seed_bid_tracker_from_workbook')]
        for p in self.patches:
            p.start()
        app.init_db()
        with closing(app.db()) as con, con:
            con.execute("INSERT INTO purchase_orders (id,po_number,job_number,customer_name,vendor,description,status,created_at,updated_at) VALUES (100,'PO-TEST','204','Example','Test','Test','Issued','2026-10-27','2026-10-27')")
            con.execute("INSERT INTO purchase_order_invoices (id,po_id,invoice_file,invoice_amount,uploaded_at) VALUES (100,100,'invoice.pdf',10,'2026-10-27T12:00:00')")
            con.execute("INSERT INTO fieldwise_parts_catalogs (id,source_file,uploaded_at) VALUES (100,'test.xlsx','2026-10-27')")
            con.execute("INSERT INTO fieldwise_parts (catalog_id,tablet_label,part_number,invoice_description) VALUES (100,'Widget DSG 128','PART','Widget')")
        (app.UPLOAD_DIR / 'invoice.pdf').write_bytes(b'%PDF-test-source')
        self.parsed = {'vendor': 'Test', 'invoice_number': 'INV1', 'date': '2026-10-27',
                       'parse_status': 'Reconciled lines',
                       'items': [{'item': 'PART', 'description': 'Widget', 'line_number': 1,
                                  'qty': 2, 'unit': 'ea', 'unit_cost': 5, 'amount': 10}]}
        self.parser = patch.object(app, 'import_vendor_invoice_pdf', return_value=self.parsed)
        self.parser.start()
        self.actor = {'id': 1, 'role': 'User', 'username': 'reviewer@example.com'}
        self.permissions = {'po_review': {'can_view': 1, 'can_edit': 1}}
        self.user_patch = patch.object(app, 'current_user', side_effect=lambda _: self.actor)
        self.role_patch = patch.object(app, 'role_permissions', side_effect=lambda _: self.permissions)
        self.user_patch.start()
        self.role_patch.start()
        with closing(review.store(app)) as c:
            review.init_schema(c)
            with c:
                rec.import_sales_rows(c, [dict(ticket_number='1', ticket_date='2026-10-27',
                    order_number='204', customer='Example', item='Widget DSG 128', description='Widget',
                    quantity=2, rate=10, amount=20, status='Submitted')],
                    '2026-10-27', 'test.xlsx', 'one', 'test')
        self.server = ThreadingHTTPServer(('127.0.0.1', 0), app.Handler)
        self.thread = threading.Thread(target=self.server.serve_forever, daemon=True)
        self.thread.start()
        self.url = 'http://127.0.0.1:' + str(self.server.server_port)

    def tearDown(self):
        self.server.shutdown()
        self.server.server_close()
        self.thread.join()
        self.user_patch.stop()
        self.role_patch.stop()
        self.parser.stop()
        for p in reversed(self.patches):
            p.stop()
        gc.collect()
        self.tmp.cleanup()

    def request(self, path='/api/material-review', payload=None, origin=None):
        headers = {'Content-Type': 'application/json'}
        if origin:
            headers['Origin'] = origin
        request = Request(self.url + path, data=json.dumps(payload).encode() if payload is not None else None,
                          headers=headers)
        try:
            with urlopen(request, timeout=10) as response:
                raw = response.read()
                return response.status, json.loads(raw) if 'application/json' in response.headers.get('Content-Type', '') else raw.decode()
        except HTTPError as exc:
            return exc.code, json.loads(exc.read())

    def test_main_dashboard_navigation_and_local_data_refresh(self):
        status, html = self.request('/')
        self.assertEqual(status, 200)
        self.assertIn('data-tab="boughtVsSold"', html)
        self.assertNotIn('boughtVsSoldFrame', html)
        self.assertIn('id="mr-count"', html)
        self.assertIn('window.loadBoughtVsSoldReview', html)
        self.assertIn("target.closest('#boughtVsSold')", html)
        self.assertIn("boughtVsSold: 'po_review'", html)
        self.assertIn("'closedPo','boughtVsSold'", html)
        self.assertNotIn('poInvoiceAudit', html)
        self.assertNotIn('/api/po-invoice-audit', html)
        status, state = self.request()
        self.assertEqual(status, 200)
        self.assertEqual(state['source_invoice_count'], 1)
        self.assertEqual(len(state['purchases']), 1)
        self.assertTrue(state['can_edit'])
        self.assertEqual(self.request('/bought-vs-sold')[0], 200)
        self.assertIn('How this works', self.request('/bought-vs-sold')[1])
        self.assertEqual(self.request('/bought-vs-sold-guide')[0], 200)
        # A new PO attachment is discovered without changing a snapshot JSON.
        with app.db() as c:
            c.execute("INSERT INTO purchase_order_invoices (po_id,invoice_file,invoice_amount,uploaded_at) VALUES (100,'second.pdf',10,'2026-10-27T13:00:00')")
        (app.UPLOAD_DIR / 'second.pdf').write_bytes(b'%PDF-second-source')
        self.assertEqual(self.request()[1]['source_invoice_count'], 2)
        # Parsing is read-only: no extra cost records were inserted.
        self.assertEqual(app.one('SELECT COUNT(*) AS n FROM cost_records')['n'], 0)

    def test_issue_persists_with_group_status_history_and_conflict_protection(self):
        buy = self.request()[1]['purchases'][0]
        payload = dict(purchase_key=buy['purchase_key'], expected_revision=0,
                       issue_group='Missing on ticket', description='Check purchased quantity',
                       status='Open', notes='Investigating')
        self.assertEqual(self.request('/api/material-review/issue', payload)[0], 200)
        self.assertEqual(self.request('/api/material-review/issue', payload)[0], 409)
        state = self.request()[1]
        self.assertEqual(state['purchases'][0]['timing_state'], 'Open investigation')
        self.assertEqual(state['purchases'][0]['issue']['issue_group'], 'Missing on ticket')
        payload.update(expected_revision=1, status='Resolved', notes='Checked with office')
        self.assertEqual(self.request('/api/material-review/issue', payload)[0], 200)
        state = self.request()[1]
        self.assertEqual(state['purchases'][0]['confirmed_qty'], 0)
        self.assertEqual(len(state['events']), 2)
        with closing(review.store(app)) as c:
            self.assertEqual(c.execute('SELECT revision FROM material_line_issues').fetchone()[0], 2)

    def test_confirm_and_reverse_quantity_keep_history(self):
        buy = self.request()[1]['purchases'][0]
        payload = dict(purchase_key=buy['purchase_key'], sales_line_id=buy['candidates'][0]['sales_line_id'],
                       bought_qty=2, bought_units_per_sale_unit=1, unit_note='1 EA purchased = 1 EA sold')
        status, result = self.request('/api/material-review/confirm', payload)
        self.assertEqual(status, 200)
        self.assertEqual(self.request()[1]['purchases'][0]['confirmed_qty'], 2)
        self.assertEqual(self.request('/api/material-review/confirm', payload)[0], 400)
        self.assertEqual(self.request('/api/material-review/unconfirm', {'id': result['id']})[0], 200)
        state = self.request()[1]
        self.assertEqual(state['purchases'][0]['confirmed_qty'], 0)
        self.assertEqual(len(state['events']), 2)

    def test_source_correction_invalidates_confirmation_but_retains_issue(self):
        buy = self.request()[1]['purchases'][0]
        self.request('/api/material-review/issue', dict(purchase_key=buy['purchase_key'], expected_revision=0,
            issue_group='Check source', description='Verify invoice item', status='Open', notes=''))
        self.request('/api/material-review/confirm', dict(purchase_key=buy['purchase_key'],
            sales_line_id=buy['candidates'][0]['sales_line_id'], bought_qty=2,
            bought_units_per_sale_unit=1, unit_note='EA to EA'))
        (app.UPLOAD_DIR / 'invoice.pdf').write_bytes(b'%PDF-source-corrected')
        state = self.request()[1]
        self.assertEqual(state['purchases'][0]['confirmed_qty'], 0)
        self.assertEqual(len(state['orphan_issues']), 1)
        self.assertEqual(len(state['events']), 2)
        # The investigation can still be closed after replacement, without
        # transferring a confirmation onto the new document.
        self.assertEqual(self.request('/api/material-review/issue', dict(purchase_key=buy['purchase_key'],
            expected_revision=1, issue_group='Check source', description='Original document replaced',
            status='Resolved', notes='Replacement requires its own quantity check'))[0], 200)
        self.assertEqual(self.request()[1]['purchases'][0]['confirmed_qty'], 0)

    def test_read_only_and_field_users_cannot_change_review(self):
        self.permissions['po_review']['can_edit'] = 0
        self.assertFalse(self.request()[1]['can_edit'])
        self.assertEqual(self.request('/bought-vs-sold-guide')[0], 200)
        for action in ('issue', 'confirm', 'unconfirm', 'upload'):
            self.assertEqual(self.request('/api/material-review/' + action, {})[0], 403)
        self.permissions = {}
        self.actor['role'] = 'Field PO'
        self.assertEqual(self.request()[0], 403)
        self.assertEqual(self.request('/bought-vs-sold')[0], 403)
        self.assertEqual(self.request('/bought-vs-sold-guide')[0], 403)

    def test_cross_origin_writes_rejected(self):
        self.assertEqual(self.request('/api/material-review/issue', {}, origin='http://unrelated.example')[0], 403)

    def test_invoice_parsing_does_not_lock_other_po_writes(self):
        with closing(app.db()) as con, con:
            con.execute("INSERT INTO purchase_order_invoices (po_id,invoice_file,invoice_amount,uploaded_at) VALUES (100,'second.pdf',10,'2026-10-27T13:00:00')")
        (app.UPLOAD_DIR / 'second.pdf').write_bytes(b'%PDF-second-source')
        def parse_while_office_edits(*args, **kwargs):
            with closing(app.db()) as writer, writer:
                writer.execute('PRAGMA busy_timeout=100')
                writer.execute("UPDATE purchase_orders SET description='Office edit during parsing' WHERE id=100")
            return self.parsed
        with closing(review.store(app)) as con:
            with patch.object(app, 'import_vendor_invoice_pdf', side_effect=parse_while_office_edits):
                invoices, docs = review.invoice_sources(app, con)
            self.assertEqual(len(invoices), 2)
            self.assertEqual(len(docs), 2)
            con.commit()

    def test_retired_audit_is_not_available(self):
        for path in ('/api/po-invoice-audit', '/api/po-invoice-audit-issues', '/api/po-invoice-audit-omissions'):
            self.assertEqual(self.request(path)[0], 403)
            self.assertEqual(self.request(path, {})[0], 403)
        with closing(app.db()) as con:
            self.assertEqual(con.execute("SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name LIKE 'po_invoice_audit_%'").fetchone()[0], 0)
        self.assertFalse(hasattr(app, 'po_invoice_audit_result'))
        self.assertEqual(self.request()[0], 200)

    def test_daily_upload_endpoint_preserves_overlap_without_double_counting(self):
        import openpyxl
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(['Ticket Number', 'Ticket Date', 'Order Number', 'Customer Name', 'Item',
                   'Description', 'Quantity', 'Rate', 'Sub Total', 'Status'])
        ws.append(['1', '2026-10-27', '204', 'Example', 'Widget DSG 128', 'Widget', 2, 10, 20, 'Submitted'])
        buffer = io.BytesIO()
        wb.save(buffer)
        wb.close()
        boundary = 'material-review-test-boundary'
        body = (f'--{boundary}\r\nContent-Disposition: form-data; name="window_end"\r\n\r\n2026-10-27\r\n'
                f'--{boundary}\r\nContent-Disposition: form-data; name="file"; filename="daily.xlsx"\r\n'
                'Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\r\n\r\n').encode()
        body += buffer.getvalue() + f'\r\n--{boundary}--\r\n'.encode()
        results = []
        for _ in range(2):
            req = Request(self.url + '/api/material-review/upload', data=body,
                          headers={'Content-Type': 'multipart/form-data; boundary=' + boundary})
            with urlopen(req, timeout=10) as response:
                results.append(json.loads(response.read()))
        self.assertFalse(results[0]['duplicate_file'])
        self.assertTrue(results[1]['duplicate_file'])
        self.assertEqual(len(self.request()[1]['sales']), 1)

    def test_ticket_return_requires_review_and_invalidates_confirmation(self):
        buy = self.request()[1]['purchases'][0]
        self.request('/api/material-review/confirm', dict(purchase_key=buy['purchase_key'],
            sales_line_id=buy['candidates'][0]['sales_line_id'], bought_qty=2,
            bought_units_per_sale_unit=1, unit_note='EA to EA'))
        original = self.request()[1]['sales'][0]
        returned = {**original, 'ticket_number': '2', 'quantity': -2, 'amount': -20}
        for r in (original, returned):
            for k in ('sales_line_id', 'ticket_revision', 'counts_as_sold'):
                r.pop(k, None)
        with closing(review.store(app)) as c, c:
            rec.import_sales_rows(c, [original, returned], '2026-10-27', 'revised.xlsx', 'credit', 'test')
        state = self.request()[1]
        self.assertEqual(state['purchases'][0]['base_state'], 'Ticket credit / return — review')
        self.assertEqual(state['purchases'][0]['confirmed_qty'], 0)
        self.assertEqual(len(state['stale_matches']), 1)


if __name__ == '__main__':
    unittest.main()
