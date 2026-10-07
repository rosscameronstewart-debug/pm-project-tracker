import json
import unittest
from datetime import datetime, timedelta
from urllib.request import Request, urlopen
from urllib.error import HTTPError
from unittest.mock import patch
from types import SimpleNamespace

import app
import session_security as security
import test_access_control as access_tests


class SessionTests(unittest.TestCase):
    setUpClass = classmethod(access_tests.AccessTests.setUpClass.__func__)
    tearDownClass = classmethod(access_tests.AccessTests.tearDownClass.__func__)

    def setUp(self):
        type(self).user_patch.stop()
        app.execute('UPDATE users SET active=1,must_change_password=0,password_hash=? WHERE id=101', (app.hash_password('OldPassword123!'),))

    def tearDown(self):
        app.execute('DELETE FROM user_sessions')
        type(self).user_patch.start()

    def request(self, path, token=None, method='GET', data=None, secure=False):
        headers={'Content-Type':'application/json'}
        if token:
            headers['Cookie']=('__Host-pm_session=' if secure else 'pm_session=')+token
        if secure:
            headers.update({'Host':'dashboard.example.com','X-Forwarded-Proto':'https'})
        try:
            body=(json.dumps(data).encode() if data is not None else b'') if method!='GET' else None
            with urlopen(Request(self.url+path,data=body,method=method,headers=headers),timeout=10) as response:
                return response.status,response.read(),response.headers
        except HTTPError as error:
            return error.code,error.read(),error.headers

    def test_temporary_password_cannot_bypass_prompt(self):
        app.execute('UPDATE users SET must_change_password=1 WHERE id=101')
        token=app.create_session(101)
        self.assertEqual(self.request('/api/me',token)[0],200)
        for path,method in (('/api/job-order-report','GET'),('/api/purchase-orders','POST'),
                            ('/api/projects/1','PUT'),('/api/subprojects/1','DELETE')):
            status,body,_=self.request(path,token,method)
            self.assertEqual(status,403)
            self.assertEqual(json.loads(body)['code'],'password_change_required')
        old_other=app.create_session(101)
        status,_,headers=self.request('/api/change-password',token,'POST',dict(current_password='OldPassword123!',new_password='NewPassword123!',confirm_password='NewPassword123!'))
        self.assertEqual(status,200)
        fresh=headers['Set-Cookie'].split(';')[0].split('=',1)[1]
        self.assertNotEqual(fresh,token)
        self.assertEqual(self.request('/api/me',old_other)[0],401)
        self.assertEqual(self.request('/api/me',token)[0],401)
        self.assertEqual(self.request('/api/me',fresh)[0],200)
        self.assertEqual(app.one('SELECT must_change_password FROM users WHERE id=101')['must_change_password'],0)

    def test_admin_reset_revokes_legacy_and_https_sessions(self):
        admin=app.create_session(1)
        app.create_session(101)
        app.create_session(101,secure=True)
        self.assertEqual(self.request('/api/users/101',admin,'PUT',{'password':'ResetPassword123!'})[0],200)
        self.assertEqual(app.one('SELECT count(*) AS n FROM user_sessions WHERE user_id=101')['n'],0)
        self.assertEqual(app.one('SELECT must_change_password FROM users WHERE id=101')['must_change_password'],1)

    def test_idle_and_absolute_expiry(self):
        token=app.create_session(101)
        now=datetime.now()
        app.execute('UPDATE user_sessions SET created_at=?,expires_at=? WHERE session_token=?',
                    ((now-timedelta(hours=13)).isoformat(),(now+timedelta(minutes=20)).isoformat(),token))
        self.assertEqual(self.request('/api/me',token)[0],401)
        app.execute('UPDATE user_sessions SET created_at=?,expires_at=? WHERE session_token=?',
                    (now.isoformat(),(now-timedelta(minutes=1)).isoformat(),token))
        self.assertEqual(self.request('/api/me',token)[0],401)
        created=now-timedelta(hours=11,minutes=50)
        app.execute('UPDATE user_sessions SET created_at=?,expires_at=? WHERE session_token=?',
                    (created.isoformat(),(now+timedelta(minutes=5)).isoformat(),token))
        self.assertEqual(self.request('/api/me',token)[0],200)
        refreshed=datetime.fromisoformat(app.one('SELECT expires_at FROM user_sessions WHERE session_token=?',(token,))['expires_at'])
        self.assertLessEqual(refreshed,created+timedelta(hours=12))

    def test_https_and_legacy_can_coexist_without_cookie_downgrade(self):
        with patch.object(security,'HTTPS_HOSTS',{'dashboard.example.com'}),patch.object(security,'TRUSTED_PROXY_IPS',{'127.0.0.1'}):
            payload={'username':'field1@example.com','password':'OldPassword123!'}
            status,_,headers=self.request('/api/login',method='POST',data=payload)
            self.assertEqual(status,200)
            self.assertNotIn('Secure',headers['Set-Cookie'])
            legacy=headers['Set-Cookie'].split(';')[0].split('=',1)[1]
            status,_,headers=self.request('/api/login',method='POST',data=payload,secure=True)
            self.assertEqual(status,200)
            self.assertIn('__Host-pm_session=',headers['Set-Cookie'])
            self.assertIn('; Secure',headers['Set-Cookie'])
            https=headers['Set-Cookie'].split(';')[0].split('=',1)[1]
            self.assertEqual(self.request('/api/me',legacy)[0],200)
            self.assertEqual(self.request('/api/me',https,secure=True)[0],200)
            self.assertEqual(self.request('/api/me',https)[0],401)
            self.assertEqual(self.request('/api/me',legacy,secure=True)[0],401)
            status,_,headers=self.request('/api/logout',https,'POST',secure=True)
            self.assertEqual(status,200)
            self.assertIn('; Secure',headers['Set-Cookie'])
            self.assertIn('Max-Age=0',headers['Set-Cookie'])
            fake=SimpleNamespace(headers={'Host':'dashboard.example.com','X-Forwarded-Proto':'https'},client_address=('203.0.113.1',1))
            self.assertFalse(security.secure_request(fake))


if __name__=='__main__':
    unittest.main()
