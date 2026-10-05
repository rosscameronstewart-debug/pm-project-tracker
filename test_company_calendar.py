import json
import gc
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
import company_calendar


class CalendarTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.path = Path(self.tmp.name) / 'calendar.sqlite3'
        with sqlite3.connect(self.path) as con:
            con.executescript("CREATE TABLE projects(id INTEGER PRIMARY KEY, name TEXT, project_code TEXT); CREATE TABLE users(id INTEGER PRIMARY KEY); INSERT INTO users VALUES(1); INSERT INTO projects VALUES(1,'Example','P1');")
            company_calendar.initialize(con)
        self.actor = {'id': 1, 'role': 'Admin', 'username': company_calendar.PREVIEW_USERNAME}
        self.data = dict(title='Start job', category='Project start', status='Scheduled',
                         start_date='2026-10-05', end_date='2026-10-07', project_id=1,
                         crew_count=5, leadership_hours=2)

    def tearDown(self):
        gc.collect()
        self.tmp.cleanup()

    def test_persistence_edit_and_project_deletion(self):
        with sqlite3.connect(self.path) as con:
            con.execute('PRAGMA foreign_keys=ON')
            event_id = company_calendar.save(con, self.data, self.actor)
            company_calendar.save(con, dict(self.data, id=event_id, status='Complete'), self.actor)
            con.execute('DELETE FROM projects WHERE id=1')
            self.assertEqual(con.execute('SELECT status,project_id,crew_count FROM company_calendar_events').fetchone(), ('Complete', None, 5))
        with sqlite3.connect(self.path) as con:
            self.assertEqual(con.execute('SELECT count(*) FROM company_calendar_events').fetchone()[0], 1)

    def test_invalid_input(self):
        cases = [dict(end_date='2026-10-01'), dict(start_date='2026-02-30'), dict(crew_count=-1),
                 dict(crew_count='2.5'), dict(leadership_hours='nan'), dict(category='invalid'),
                 dict(project_id=999), dict(title='')]
        with sqlite3.connect(self.path) as con:
            for changes in cases:
                with self.subTest(changes=changes), self.assertRaises(ValueError):
                    company_calendar.save(con, dict(self.data, **changes), self.actor)
            with self.assertRaises(LookupError):
                company_calendar.save(con, dict(self.data, id=999), self.actor)

    def test_private_preview_permissions(self):
        for role in ('Admin', 'User', 'Read Only', 'TX/Read Only', 'Field PO'):
            other = {'id':3, 'role':role, 'username':'other@example.com'}
            self.assertFalse(app.can_view_permission(other, 'company_calendar'))
            self.assertFalse(app.can_edit_permission(other, 'company_calendar'))
        self.assertTrue(app.can_view_permission(self.actor, 'company_calendar'))
        self.assertTrue(app.can_edit_permission(self.actor, 'company_calendar'))
        self.assertFalse(company_calendar.preview_allowed(dict(self.actor, active=0)))
        self.assertNotIn('company_calendar', [p['key'] for p in app.PERMISSION_DEFINITIONS])

    def test_http_permissions_and_validation(self):
        with patch.object(app, 'DB_PATH', self.path), patch.object(app, 'current_user', side_effect=lambda _: self.actor), patch.object(app, 'role_permissions', return_value={'company_calendar': {'can_view': 1, 'can_edit': 0}}), patch.object(app, 'log_activity'):
            server = ThreadingHTTPServer(('127.0.0.1', 0), app.Handler)
            thread = threading.Thread(target=server.serve_forever, daemon=True)
            thread.start()
            url = 'http://127.0.0.1:%s/api/company-calendar' % server.server_port
            def post(data):
                return urlopen(Request(url, data=json.dumps(data).encode(), headers={'Content-Type':'application/json'}))
            try:
                with post(self.data) as response:
                    self.assertTrue(json.load(response)['id'])
                with post(dict(self.data, end_date='2026-10-01')) as response:
                    self.fail('Invalid dates accepted')
            except HTTPError as error:
                self.assertEqual(error.code, 400)
            finally:
                try:
                    self.actor = {'id':1,'role':'Read Only'}
                    with self.assertRaises(HTTPError) as denied:
                        urlopen(url)
                    self.assertEqual(denied.exception.code,403)
                    with self.assertRaises(HTTPError) as denied:
                        post(self.data)
                    self.assertEqual(denied.exception.code,403)
                    self.actor = {'id':3,'role':'Admin','username':'other@example.com'}
                    with self.assertRaises(HTTPError) as denied:
                        urlopen(url)
                    self.assertEqual(denied.exception.code,403)
                    with self.assertRaises(HTTPError) as denied:
                        post(self.data)
                    self.assertEqual(denied.exception.code,403)
                    self.actor = None
                    with self.assertRaises(HTTPError) as denied:
                        urlopen(url)
                    self.assertEqual(denied.exception.code,401)
                finally:
                    server.shutdown()
                    server.server_close()
                    thread.join()


if __name__ == '__main__':
    unittest.main()
