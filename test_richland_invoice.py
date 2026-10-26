import json
from pathlib import Path
import unittest
from unittest.mock import patch

import app


class RichlandInvoiceTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.pages = json.loads((Path(__file__).parent / 'tests/fixtures/richland_377958_ocr.json').read_text())
        cls.text = app.richland_invoice_ocr_text(cls.pages)

    def test_scanned_sample_metadata_and_all_five_costs(self):
        parsed = app.parse_richland_pump_invoice(self.text)
        self.assertEqual(parsed['vendor'], 'Richland Pump & Supply')
        self.assertEqual(parsed['invoice_number'], '377958')
        self.assertEqual(parsed['invoice_date'], '2026-10-08')
        self.assertEqual(parsed['order_number'], '2026-0333')
        self.assertEqual([row['amount'] for row in parsed['item_lines']], [192.52, 55.97, 14.93, 52.28, 8.68])
        self.assertEqual([row['qty'] for row in parsed['item_lines']], [1] * 5)
        self.assertEqual(parsed['item_lines'][1]['product_code'], '590')
        self.assertEqual(parsed['item_lines'][2]['product_code'], 'Little-Joe')
        self.assertAlmostEqual(sum(row['amount'] for row in parsed['item_lines']), parsed['total_due'])

    def test_delivery_ticket_is_not_a_second_invoice(self):
        ticket = '\nDELIVERY TICKET\nQuantity Item Code Description Price Each Amount\n1 OTHER duplicate 999.00 999.00T\nSubtotal 999.00\nTotal $999.00'
        parsed = app.parse_richland_pump_invoice(self.text + ticket)
        self.assertEqual(len(parsed['item_lines']), 5)
        self.assertEqual(parsed['total_due'], 324.38)

    def test_missing_item_and_extra_charges_require_review(self):
        for text in [self.text.replace('55.97T', 'unreadable'), self.text.replace('$324.38', '$330.00')]:
            with self.subTest(text=text), self.assertRaises(RuntimeError):
                app.parse_richland_pump_invoice(text)

    def test_explicit_quantity_is_kept_and_checked(self):
        text = self.text.replace('CN1290S\t', '2\tCN1290S\t').replace('192.52T', '385.04T').replace('324.38', '516.90')
        parsed = app.parse_richland_pump_invoice(text)
        self.assertEqual(parsed['item_lines'][0]['qty'], 2)
        with self.assertRaises(RuntimeError):
            app.parse_richland_pump_invoice(text.replace('2\tCN1290S', '3\tCN1290S'))

    def test_import_routes_to_parser_and_preserves_duplicate_protection(self):
        raw = '\n'.join(cell['text'] for page in self.pages for cell in page)
        with patch.object(app, 'extract_pdf_text', return_value=raw), \
             patch.object(app.pdfplumber, 'open') as pdf, \
             patch.object(app, 'extract_pdf_ocr_pages', return_value=self.pages), \
             patch.object(app, 'one', return_value=None), patch.object(app, 'execute') as insert:
            pdf.return_value.__enter__.return_value.pages = []
            result = app.import_vendor_invoice_pdf('scan.pdf', 99)
            self.assertEqual(result['count'], 5)
            self.assertEqual(result['invoice_number'], '377958')
            self.assertEqual(result['order_number'], '2026-0333')
            self.assertAlmostEqual(sum(call.args[1][10] for call in insert.call_args_list), 324.38)
            insert.reset_mock()
            with patch.object(app, 'one', side_effect=[None, None, {'line_count': 5, 'total_amount': 324.38, 'source_file': 'scan.pdf'}]):
                duplicate = app.import_vendor_invoice_pdf('renamed-scan.pdf', 99)
            self.assertTrue(duplicate['duplicate'])
            insert.assert_not_called()


if __name__ == '__main__':
    unittest.main()
