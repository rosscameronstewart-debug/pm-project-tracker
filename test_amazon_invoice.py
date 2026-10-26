import unittest
from unittest.mock import patch

import app


SAMPLE = '''Final Details for Order #113-7177756-8388221
Order Placed: September 9, 2026
PO number : 2026-0086
Amazon.com order number: 113-7177756-8388221
Order Total: USD 49.49
Shipped on September 10, 2026
Items Ordered Price
1 of: GDSTIME AXIAL 2060 Muffin Fan, 120V AC 200mm x 60mm High Airflow Dual Ball $49.49
Sold by and invoiced on behalf of: GDSTIME (seller profile)
Business Price
Condition: New
Shipping Address: Item(s) Subtotal: $49.49
Shipping & Handling: $0.00
Total before tax: $49.49
Sales Tax: $0.00
Total for This Shipment: $49.49
Payment information
Payment Method: Item(s) Subtotal: USD 49.49
Shipping & Handling: USD 0.00
Total before tax: USD 49.49
Estimated tax to be collected: USD 0.00
Grand Total: USD 49.49
Credit Card transactions: September 10, 2026: $49.49
Conditions of Use | Privacy Notice Amazon.com, Inc.
'''


class AmazonInvoiceTests(unittest.TestCase):
    def test_sample_metadata_seller_and_cost(self):
        result = app.parse_amazon_order_invoice(SAMPLE)
        self.assertEqual(result['vendor'], 'Amazon')
        self.assertEqual(result['invoice_number'], '113-7177756-8388221')
        self.assertEqual(result['invoice_date'], '2026-09-09')
        self.assertEqual(result['order_number'], '2026-0086')
        self.assertEqual(result['total_due'], 49.49)
        self.assertEqual(len(result['item_lines']), 1)
        self.assertEqual(result['item_lines'][0]['qty'], 1)
        self.assertEqual(result['item_lines'][0]['amount'], 49.49)
        self.assertIn('Sold by: GDSTIME', result['item_lines'][0]['description'])

    def test_quantity_multiplies_unit_price_and_wrapped_description_survives(self):
        text = SAMPLE.replace('1 of:', '2 of:').replace('Dual Ball $49.49', 'Dual Ball $49.49\nwith power cord')
        text = text.replace('Order Total: USD 49.49', 'Order Total: USD 98.98').replace('Grand Total: USD 49.49', 'Grand Total: USD 98.98')
        result = app.parse_amazon_order_invoice(text)
        self.assertEqual(result['item_lines'][0]['amount'], 98.98)
        self.assertIn('with power cord', result['item_lines'][0]['description'])

    def test_shipping_tax_and_discount_are_counted_once(self):
        text = SAMPLE.replace('Shipping & Handling: USD 0.00', 'Shipping & Handling: USD 5.00')
        text = text.replace('Estimated tax to be collected: USD 0.00', 'Estimated tax to be collected: USD 3.00\nPromotion(s) Applied: USD -2.00')
        text = text.replace('Grand Total: USD 49.49', 'Grand Total: USD 55.49')
        result = app.parse_amazon_order_invoice(text)
        self.assertEqual([line['amount'] for line in result['item_lines']], [49.49, 5, 3, -2])

    def test_multiple_shipments_reconcile_to_order_not_one_shipment(self):
        second = 'Shipped on September 11, 2026\nItems Ordered Price\n1 of: Another item $10.00\nSold by: Seller\nShipping Address: Item(s) Subtotal: $10.00\nTotal for This Shipment: $10.00\n'
        text = SAMPLE.replace('Payment information', second + 'Payment information').replace('Grand Total: USD 49.49', 'Grand Total: USD 59.49')
        result = app.parse_amazon_order_invoice(text)
        self.assertEqual(len(result['item_lines']), 2)
        self.assertEqual(sum(line['amount'] for line in result['item_lines']), 59.49)

    def test_unreconciled_or_unreadable_items_block_import(self):
        for text in [SAMPLE.replace('Grand Total: USD 49.49', 'Grand Total: USD 100.00'), SAMPLE.replace('Dual Ball $49.49', 'Dual Ball unreadable')]:
            with self.subTest(text=text), self.assertRaises(RuntimeError):
                app.parse_amazon_order_invoice(text)

    def test_import_routes_and_duplicate_check_uses_order_number(self):
        with patch.object(app, 'extract_pdf_text', return_value=SAMPLE), \
             patch.object(app, 'one', return_value=None), patch.object(app, 'execute') as insert:
            result = app.import_vendor_invoice_pdf('amazon.pdf', 99)
            self.assertEqual(result['count'], 1)
            self.assertEqual(insert.call_args.args[1][10], 49.49)
            self.assertEqual(insert.call_args.args[1][3], '113-7177756-8388221')
            insert.reset_mock()
            with patch.object(app, 'one', side_effect=[None, None, {'line_count': 1, 'total_amount': 49.49, 'source_file': 'amazon.pdf'}]):
                result = app.import_vendor_invoice_pdf('renamed.pdf', 99)
            self.assertTrue(result['duplicate'])
            insert.assert_not_called()


if __name__ == '__main__':
    unittest.main()
