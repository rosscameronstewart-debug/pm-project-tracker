from pathlib import Path
import unittest
from unittest.mock import patch
import app


class CedInvoiceTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.text=(Path(__file__).parent/'tests/fixtures/ced_9463_1040004.txt').read_text()

    def preview(self,text):
        with patch.object(app,'extract_pdf_text',return_value=text), \
             patch.object(app,'one',side_effect=AssertionError('No accounting queries in preview')), \
             patch.object(app,'execute',side_effect=AssertionError('No accounting writes in preview')):
            return app.import_vendor_invoice_pdf('ced.pdf',0,preview_details=True)

    def test_new_export_with_trailing_column_and_wrapped_part_codes(self):
        r=self.preview(self.text)
        self.assertEqual((r['invoice_number'],r['date'],r['order_reference']),('9463-1040004','2026-10-28','GOGO2026531'))
        self.assertEqual([x['item'] for x in r['items']],['CULLY 55016J','CULLY 40720J','POWST PSNS1/4EG'])
        self.assertEqual([x['qty'] for x in r['items']],[200,100,100])
        self.assertEqual([x['amount'] for x in r['items']],[19.70,5.91,72])
        self.assertEqual(r['parse_status'],'Reconciled lines')
        self.assertAlmostEqual(r['total_difference'],0)
        self.assertNotIn('MERCHANDISE',r['items'][-1]['description'])

    def test_c_price_is_per_hundred_and_e_is_each(self):
        r=self.preview(self.text)
        self.assertAlmostEqual(r['items'][0]['unit_cost'],9.85/100)
        self.assertAlmostEqual(r['items'][1]['unit_cost'],5.91/100)
        self.assertAlmostEqual(r['items'][2]['unit_cost'],.72)

    def test_older_export_without_trailing_number_still_parses(self):
        r=self.preview(self.text.replace(' 0.0\n','\n').replace('200 CULLY','T 200 CULLY',1))
        self.assertEqual(len(r['items']),3)
        self.assertEqual(r['parse_status'],'Reconciled lines')
        self.assertEqual(r['line_total'],97.61)

    def test_missing_item_stays_in_review_instead_of_falsely_reconciling(self):
        r=self.preview(self.text.replace("100 POWST 1/4'' STRUT NUT 100 0.72 E 72.00 0.0",'Unreadable item'))
        self.assertEqual(r['parse_status'],'Needs parser review')
        self.assertAlmostEqual(r['total_difference'],72)


if __name__=='__main__':unittest.main()
