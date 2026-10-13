import json, unittest
from urllib.request import Request,urlopen
from urllib.error import HTTPError
import app
from test_access_control import AccessTests
class CogCodeTests(unittest.TestCase):
    setUpClass=classmethod(AccessTests.setUpClass.__func__)
    tearDownClass=classmethod(AccessTests.tearDownClass.__func__)
    def request(self,path,data,method):
        try:
            with urlopen(Request(self.url+path,data=json.dumps(data).encode(),method=method,headers={'Content-Type':'application/json'})) as r:
                return r.status,json.loads(r.read())
        except HTTPError as e: return e.code,json.loads(e.read())
    def test_code_roundtrip_and_validation(self):
        type(self).actor={'id':1,'role':'Admin','username':'admin'}
        data={'name':'Code test','code':'TEST','po_code':'007'}
        status,result=self.request('/api/cog-categories',data,'POST')
        self.assertEqual(status,200)
        key=result['id']
        self.assertEqual(app.one('SELECT po_code FROM cog_categories WHERE id=?',(key,))['po_code'],'007')
        for valid in ['', '7', 'AB', 'A07', '1234', 'aB123']:
            data['po_code']=valid
            self.assertEqual(self.request('/api/cog-categories/'+str(key),data,'PUT')[0],200)
            self.assertEqual(app.one('SELECT po_code FROM cog_categories WHERE id=?',(key,))['po_code'],valid)
        for bad in ['ABC123', 'A-7', 'A 7', '\uff11\uff12\uff13']:
            data['po_code']=bad
            self.assertEqual(self.request('/api/cog-categories/'+str(key),data,'PUT')[0],400)
        data['po_code']='a07'
        self.assertEqual(self.request('/api/cog-categories/'+str(key),data,'PUT')[0],200)
        self.assertEqual(app.one('SELECT po_code FROM cog_categories WHERE id=?',(key,))['po_code'],'a07')
if __name__=='__main__': unittest.main()
