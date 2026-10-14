import unittest
from types import SimpleNamespace
import project_materials


class MaterialTests(unittest.TestCase):
    def test_invoice_line_visibility_excludes_prices_and_preserves_quantity(self):
        def rows(sql,values=()):
            if 'SELECT id,project_code,name FROM projects' in sql:return [{'id':1,'project_code':'JOB','name':'Test'}]
            if 'JOIN purchase_orders' in sql:return [{'invoice_file':'inv.pdf','po_number':'TEST-26-0001','requested_by_username':'buyer','uploaded_by_username':'office'}]
            return [{'id':1,'item':'Coupling','description':'3/4 EMT coupling','qty':25,'vendor':'DSG','record_date':'2026-10-14','ticket_or_invoice':'INV1','source_file':'inv.pdf','job_number':'641','job_label':'Test','amount':100,'rate':4}]
        result=project_materials.list_materials(SimpleNamespace(rows=rows),1)
        self.assertEqual(len(result['items']),1)
        row=result['items'][0]
        self.assertEqual(row['qty'],25)
        self.assertEqual(row['buyer'],'buyer')
        self.assertNotIn('amount',row)
        self.assertNotIn('source_file',row)
        self.assertEqual(result['warnings'],[])


if __name__=='__main__':unittest.main()
