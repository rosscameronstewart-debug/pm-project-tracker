from pathlib import Path
import unittest
from unittest.mock import patch,MagicMock
import app


class AutomationXInvoiceTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.text=(Path(__file__).parent/'tests/fixtures/automation_x_inv445813.txt').read_text()

    def test_invoice_items_and_separate_charges_reconcile(self):
        r=app.parse_automation_x_invoice(self.text)
        self.assertEqual((r['invoice_number'],r['invoice_date'],r['order_number']),('INV445813','2026-10-28','Stock'))
        self.assertEqual([x['product_code'] for x in r['item_lines']],['-70616','-75259','Tariff','Shipping'])
        self.assertEqual([x['amount'] for x in r['item_lines']],[2200,113.83,86.22,27.79])
        self.assertAlmostEqual(sum(x['amount'] for x in r['item_lines']),2427.84)
        self.assertIn('QTERM-A7',r['item_lines'][1]['description'])
        self.assertNotIn('SPECIAL ORDER',r['item_lines'][1]['description'])

    def test_due_and_ship_dates_are_not_invoice_date_fallbacks(self):
        with self.assertRaises(RuntimeError):
            app.parse_automation_x_invoice(self.text.replace('Please Send Payments to: 10/28/2026','Please Send Payments to:'))

    def test_missing_lines_or_inconsistent_prices_fail_closed(self):
        for old,new in [('1 $113.83 0% $113.83','2 $113.83 0% $113.83'),
                        ('Subtotal: $2,400.05','Subtotal: $2,500.05'),
                        ('Total: $2,427.84','Total: $2,527.84'),
                        ('Shipping & Handling: $27.79','Shipping & Handling: unreadable'),
                        ('2 -75259 Beijer Adapter Plate 100-1157 1 $113.83 0% $113.83','Unreadable item')]:
            with self.subTest(old=old),self.assertRaises(RuntimeError):
                app.parse_automation_x_invoice(self.text.replace(old,new))

    def test_bundled_order_is_not_counted_again(self):
        text=self.text+'\nSales Order\nSO393698\nDO NOT PAY - THIS IS NOT AN INVOICE 10/9/2026\n'+self.text
        self.assertEqual(len(app.parse_automation_x_invoice(text)['item_lines']),4)

    def test_multiple_invoice_numbers_require_review(self):
        with self.assertRaises(RuntimeError):
            app.parse_automation_x_invoice(self.text+self.text.replace('INV445813','INV999999'))

    def test_preview_filters_non_invoice_pages_without_accounting_writes(self):
        pages=[]
        for text in (self.text,'Sales Order\nSO393698\nDO NOT PAY - THIS IS NOT AN INVOICE\n'+self.text,
                     'Packing Slip\n601831\n10/9/2026'):
            p=MagicMock();p.extract_text.return_value=text;pages.append(p)
        pdf=MagicMock();pdf.__enter__.return_value.pages=pages
        with patch.object(app,'extract_pdf_text',return_value='\n'.join(p.extract_text() for p in pages)), \
             patch.object(app.pdfplumber,'open',return_value=pdf), \
             patch.object(app,'one',side_effect=AssertionError('No accounting queries')), \
             patch.object(app,'execute',side_effect=AssertionError('No accounting writes')):
            r=app.import_vendor_invoice_pdf('bundle.pdf',0,preview_details=True)
        self.assertEqual(r['parse_status'],'Reconciled lines')
        self.assertEqual(r['total_difference'],0)
        self.assertEqual(len(r['items']),4)
        self.assertEqual(r['items'][0]['qty'],1)


if __name__=='__main__':unittest.main()
