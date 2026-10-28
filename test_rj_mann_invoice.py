from pathlib import Path
import unittest
from unittest.mock import patch

import app


class RJMannInvoiceTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.text = (Path(__file__).parent / 'tests/fixtures/rj_mann_ari195562.txt').read_text()

    def test_sample_metadata_and_item_math(self):
        parsed = app.parse_rj_mann_invoice(self.text)
        self.assertEqual(parsed['vendor'], 'RJ Mann & Associates, Inc.')
        self.assertEqual(parsed['invoice_number'], 'ARI195562')
        self.assertEqual(parsed['invoice_date'], '2026-10-26')
        self.assertEqual(parsed['order_number'], 'COR260500')
        self.assertEqual(len(parsed['item_lines']), 1)
        item = parsed['item_lines'][0]
        self.assertEqual(item['product_code'], 'MUROPLFCA1000OS')
        self.assertEqual(item['description'], 'PRESSURE SWITCHGAUGE')
        self.assertEqual(item['qty'], 3)
        self.assertEqual(item['unit_price'], 349.59)
        self.assertEqual(item['amount'], 1048.77)

    def test_due_date_is_never_used_as_invoice_date(self):
        with self.assertRaises(RuntimeError):
            app.parse_rj_mann_invoice(self.text.replace('Date: 10/26/2026', 'Invoice date unreadable'))

    def test_payment_does_not_reduce_purchase_cost(self):
        text = self.text.replace('Payment/Credit Applied 0.00', 'Payment/Credit Applied 100.00').replace('Total (USD): 1,048.77', 'Total (USD): 948.77')
        self.assertEqual(app.parse_rj_mann_invoice(text)['total_due'], 1048.77)

    def test_missing_or_misread_lines_require_review(self):
        for text in (self.text.replace('3.00 MURO', '2.00 MURO'),
                     self.text.replace('3.00 MUROPLFCA1000OS PRESSURE SWITCHGAUGE 349.59 1,048.77', 'Unreadable item'),
                     self.text.replace('Total Invoice Amount 1,048.77', 'Total Invoice Amount 1,148.77')):
            with self.subTest(text=text), self.assertRaises(RuntimeError):
                app.parse_rj_mann_invoice(text)

    def test_multiple_lines_and_tax_kept_separate(self):
        text = self.text.replace('ATTENTION:', '2.00 PART-2 SECOND ITEM 10.00 20.00\nATTENTION:')
        text = text.replace('Subtotal 1,048.77', 'Subtotal 1,068.77').replace('Tax 0.00', 'Tax 5.00')
        text = text.replace('Total Invoice Amount 1,048.77', 'Total Invoice Amount 1,073.77')
        parsed = app.parse_rj_mann_invoice(text)
        self.assertEqual(len(parsed['item_lines']), 3)
        self.assertEqual(parsed['item_lines'][-1]['product_code'], 'Sales Tax')
        self.assertAlmostEqual(sum(r['amount'] for r in parsed['item_lines']), 1073.77)

    def test_preview_routes_to_parser_without_importing_cost_records(self):
        with patch.object(app, 'extract_pdf_text', return_value=self.text), \
             patch.object(app, 'one', side_effect=AssertionError('Preview must not query import targets')), \
             patch.object(app, 'execute', side_effect=AssertionError('Preview must not write')):
            preview = app.import_vendor_invoice_pdf('invoice.pdf', 0, preview_details=True)
        self.assertEqual(preview['date'], '2026-10-26')
        self.assertEqual(preview['parse_status'], 'Reconciled lines')
        self.assertEqual(preview['items'][0]['qty'], 3)
        self.assertEqual(preview['items'][0]['unit_cost'], 349.59)
        self.assertEqual(preview['total_difference'], 0)


if __name__ == '__main__':
    unittest.main()
