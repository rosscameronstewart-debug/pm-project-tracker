from pathlib import Path
import unittest
from unittest.mock import patch
import app


class EagleInvoiceTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.text=(Path(__file__).parent/'tests/fixtures/eagle_inv569680.txt').read_text()

    def test_four_material_lines_plus_separate_tax_reconcile(self):
        r=app.parse_eagle_automation_invoice(self.text)
        self.assertEqual((r['invoice_number'],r['invoice_date'],r['order_number']),('INV569680','2026-10-28','HIB-26-0451'))
        self.assertEqual([x['product_code'] for x in r['item_lines']],['Z9-P','VALU890-12-N4','LMR-195-NM-SMAM-36','LMR-400-NM-NM-180','Sales Tax'])
        self.assertEqual([x['amount'] for x in r['item_lines']],[1003.85,235.35,55.97,51.52,111.10])
        self.assertAlmostEqual(sum(x['amount'] for x in r['item_lines']),1457.79)
        self.assertIn('402-66-6987',r['item_lines'][0]['description'])
        self.assertIn('12dbi gain',r['item_lines'][1]['description'])
        self.assertNotIn('Thank you',r['item_lines'][-2]['description'])

    def test_missing_invoice_date_does_not_use_due_date(self):
        with self.assertRaises(RuntimeError):
            app.parse_eagle_automation_invoice(self.text.replace('10/28/2026 INV569680','INV569680'))

    def test_incomplete_lines_or_inconsistent_totals_require_review(self):
        for old,new in [('1,003.85 1,003.85','1,003.85 1,003.95'),
                        ('Subtotal 1,346.69','Subtotal 1,446.69'),
                        ('Total $1,457.79','Total $1,557.79'),
                        ('1 VALU890-12-N4 Antenna, WaveLink, 902-928 MHz, 235.35 235.35','Unreadable item')]:
            with self.subTest(old=old), self.assertRaises(RuntimeError):
                app.parse_eagle_automation_invoice(self.text.replace(old,new))

    def test_shipping_is_separate_and_not_a_material_quantity(self):
        t=self.text.replace('Pick-Up) 0.00','Pick-Up) 25.00').replace('Total $1,457.79','Total $1,482.79')
        r=app.parse_eagle_automation_invoice(t)
        shipping=next(x for x in r['item_lines'] if x['product_code']=='Shipping')
        self.assertEqual(shipping['amount'],25)

    def test_preview_reads_prices_without_writing_cost_or_sales_records(self):
        with patch.object(app,'extract_pdf_text',return_value=self.text), \
             patch.object(app,'one',side_effect=AssertionError('No import target queries in preview')), \
             patch.object(app,'execute',side_effect=AssertionError('No writes in preview')):
            r=app.import_vendor_invoice_pdf('invoice.pdf',0,preview_details=True)
        self.assertEqual(r['parse_status'],'Reconciled lines')
        self.assertEqual(r['total_difference'],0)
        self.assertEqual(r['items'][0]['qty'],1)
        self.assertEqual(r['items'][0]['unit_cost'],1003.85)
        self.assertEqual(r['items'][0]['unit'],'')  # Invoice does not specify a UOM.


if __name__=='__main__':
    unittest.main()
