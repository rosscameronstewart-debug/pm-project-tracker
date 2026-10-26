import sqlite3
import unittest
from unittest.mock import patch

import app


class PoInvoiceUpdateTests(unittest.TestCase):
    def setUp(self):
        self.con = sqlite3.connect(':memory:')
        self.con.row_factory = sqlite3.Row
        self.addCleanup(self.con.close)
        self.con.executescript('''
            CREATE TABLE purchase_orders (
                id INTEGER PRIMARY KEY, po_number TEXT, invoice_file TEXT,
                invoice_uploaded_at TEXT, invoice_uploaded_by_user_id INTEGER,
                invoice_uploaded_by_username TEXT, invoice_cost_record_id INTEGER,
                updated_at TEXT
            );
            CREATE TABLE purchase_order_invoices (
                id INTEGER PRIMARY KEY, po_id INTEGER, invoice_file TEXT,
                invoice_amount REAL, cost_record_id INTEGER, uploaded_at TEXT,
                uploaded_by_user_id INTEGER, uploaded_by_username TEXT
            );
            CREATE TABLE cost_records (id INTEGER PRIMARY KEY, raw_cost_source TEXT);
            INSERT INTO purchase_orders(id, po_number) VALUES (1, 'TEST-PO');
            INSERT INTO purchase_order_invoices VALUES
                (2, 1, 'invoice.pdf', 100, 3, '2026-10-23', 4, 'office@example.com');
            INSERT INTO cost_records VALUES (3, 'PO invoice upload');
        ''')
        self.handler = app.Handler.__new__(app.Handler)
        self.handler.path = '/api/purchase-order-invoices/2'
        for target, kwargs in [
            ('access_control.authorize', {'return_value': True}),
            ('current_user', {'return_value': {'id': 4, 'username': 'office@example.com'}}),
            ('require_editor', {'return_value': True}),
            ('can_edit_permission', {'return_value': True}),
            ('db', {'return_value': self.con}),
            ('log_activity', {}),
        ]:
            mock = patch('app.' + target, **kwargs)
            mock.start()
            self.addCleanup(mock.stop)

    def update(self, amount):
        with patch.object(app, 'parse_json', return_value={'invoice_amount': amount}), \
             patch.object(app, 'json_response', side_effect=lambda _, payload, status=200: (status, payload)):
            return self.handler.do_PUT()

    def test_put_updates_existing_invoice_and_po_cost_link(self):
        with patch.object(app, 'sync_po_invoice_cost_record', return_value=3) as sync:
            status, payload = self.update('414.25')
        self.assertEqual(status, 200)
        self.assertTrue(payload['ok'])
        self.assertEqual(sync.call_args.args[4], 414.25)
        invoice = self.con.execute('SELECT * FROM purchase_order_invoices WHERE id=2').fetchone()
        self.assertEqual(invoice['invoice_amount'], 414.25)
        po = self.con.execute('SELECT * FROM purchase_orders WHERE id=1').fetchone()
        self.assertEqual(po['invoice_cost_record_id'], 3)
        self.assertEqual(po['invoice_file'], 'invoice.pdf')

    def test_zero_amount_removes_linked_invoice_cost(self):
        status, _ = self.update('0')
        self.assertEqual(status, 200)
        self.assertEqual(self.con.execute('SELECT COUNT(*) FROM cost_records').fetchone()[0], 0)
        self.assertIsNone(self.con.execute('SELECT cost_record_id FROM purchase_order_invoices').fetchone()[0])

    def test_missing_invoice_returns_specific_error(self):
        self.handler.path = '/api/purchase-order-invoices/999'
        self.assertEqual(self.update('414.25'), (404, {'error': 'PO invoice not found.'}))

    def test_office_permission_is_required(self):
        with patch.object(app, 'can_edit_permission', return_value=False):
            status, _ = self.update('414.25')
        self.assertEqual(status, 403)
        self.assertEqual(self.con.execute('SELECT invoice_amount FROM purchase_order_invoices').fetchone()[0], 100)


if __name__ == '__main__':
    unittest.main()
