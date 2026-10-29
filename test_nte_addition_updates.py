import sqlite3
import unittest
from unittest.mock import patch

import app


class NteAdditionUpdateTests(unittest.TestCase):
    def setUp(self):
        self.con = sqlite3.connect(':memory:')
        self.con.row_factory = sqlite3.Row
        self.addCleanup(self.con.close)
        self.con.executescript('''
            CREATE TABLE nte_bucket_additions (
                id INTEGER PRIMARY KEY, bucket_id INTEGER, amount REAL,
                labor_amount REAL, material_amount REAL, equipment_amount REAL,
                status TEXT, signed_file TEXT, support_file TEXT, created_at TEXT, reason TEXT
            );
            INSERT INTO nte_bucket_additions VALUES
                (1, 2, -771912, -51912, -720000, 0, 'Pending', '', 'backup.pdf', '2026-10-26', 'Bucket Decrease');
        ''')
        self.handler = app.Handler.__new__(app.Handler)
        self.handler.path = '/api/nte-additions/1'
        for target, value in [
            ('access_control.authorize', True),
            ('current_user', {'id': 1, 'username': 'test'}),
            ('require_editor', True),
            ('can_edit_permission', True),
            ('db', self.con),
        ]:
            mock = patch('app.' + target, return_value=value)
            mock.start()
            self.addCleanup(mock.stop)
        self.log = patch('app.log_activity').start()
        self.addCleanup(patch.stopall)

    def update(self, **overrides):
        data = dict(labor_amount='-5191.20', material_amount='-7200', equipment_amount='0')
        data.update(overrides)
        with patch.object(app, 'parse_json', return_value=data), patch.object(
            app, 'json_response', side_effect=lambda _, payload, status=200: (status, payload)
        ):
            return self.handler.do_PUT()

    def test_corrects_existing_row_and_recalculates_total_with_audit(self):
        self.assertEqual(self.update()[0], 200)
        row = self.con.execute('SELECT * FROM nte_bucket_additions').fetchone()
        self.assertEqual(row['amount'], -12391.20)
        self.assertEqual(row['material_amount'], -7200)
        self.assertEqual(row['support_file'], 'backup.pdf')
        self.assertEqual(row['created_at'], '2026-10-26')
        self.assertEqual(self.con.execute('SELECT COUNT(*) FROM nte_bucket_additions').fetchone()[0], 1)
        audit = self.log.call_args.args[4]
        self.assertEqual(audit['before']['amount'], -771912)
        self.assertEqual(audit['after']['amount'], -12391.20)

    def test_approved_or_signed_records_cannot_be_changed(self):
        for status, signed in [('Approved', ''), ('Pending', 'signed.pdf')]:
            self.con.execute('UPDATE nte_bucket_additions SET status=?, signed_file=?', (status, signed))
            self.assertEqual(self.update()[0], 409)
        self.assertEqual(self.con.execute('SELECT amount FROM nte_bucket_additions').fetchone()[0], -771912)

    def test_permission_required(self):
        with patch.object(app, 'can_edit_permission', return_value=False):
            self.assertEqual(self.update()[0], 403)

    def test_reason_only_edit_preserves_amounts_and_records_before_after(self):
        with patch.object(app, 'parse_json', return_value={'reason':' Corrected scope explanation '}), patch.object(
            app, 'json_response', side_effect=lambda _, payload, status=200: (status, payload)
        ):
            self.assertEqual(self.handler.do_PUT()[0], 200)
        row=self.con.execute('SELECT * FROM nte_bucket_additions').fetchone()
        self.assertEqual(row['reason'],'Corrected scope explanation')
        self.assertEqual(row['amount'],-771912)
        self.assertEqual(row['material_amount'],-720000)
        self.assertEqual(self.log.call_args.args[4]['before'],{'reason':'Bucket Decrease'})
        self.assertEqual(self.log.call_args.args[4]['after'],{'reason':'Corrected scope explanation'})

    def test_reason_and_amounts_save_together(self):
        self.assertEqual(self.update(reason='Updated explanation')[0],200)
        row=self.con.execute('SELECT reason,amount FROM nte_bucket_additions').fetchone()
        self.assertEqual(row['reason'],'Updated explanation')
        self.assertEqual(row['amount'],-12391.20)

    def test_reason_is_locked_after_approval_or_signature(self):
        for status,signed in [('Approved',''),('Pending','signed.pdf'),('Rejected','')]:
            self.con.execute('UPDATE nte_bucket_additions SET status=?,signed_file=?',(status,signed))
            self.assertEqual(self.update(reason='Changed')[0],409)
            self.assertEqual(self.con.execute('SELECT reason FROM nte_bucket_additions').fetchone()[0],'Bucket Decrease')

    def test_invalid_reason_does_not_change_amounts(self):
        self.assertEqual(self.update(reason={'unexpected':'object'})[0],400)
        self.assertEqual(self.con.execute('SELECT amount FROM nte_bucket_additions').fetchone()[0],-771912)

    def test_invalid_or_nonfinite_amounts_rejected(self):
        for value in ['bad', 'NaN', 'Infinity', None]:
            with self.subTest(value=value):
                self.assertEqual(self.update(labor_amount=value)[0], 400)
        self.log.assert_not_called()

    def test_zero_total_rejected(self):
        self.assertEqual(self.update(labor_amount='0', material_amount='0')[0], 400)

    def test_missing_record(self):
        self.handler.path = '/api/nte-additions/99'
        self.assertEqual(self.update()[0], 404)


if __name__ == '__main__':
    unittest.main()
