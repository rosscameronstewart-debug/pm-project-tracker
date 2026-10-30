from pathlib import Path
import unittest
from unittest.mock import patch,MagicMock
import app


class PowerMationBorderTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        root=Path(__file__).parent/'tests/fixtures'
        cls.pm=(root/'powermation_30658853.txt').read_text()
        cls.bs=(root/'border_states_933407913.txt').read_text()

    def tables(self):
        return [[['CUSTOMER PO NO.','ORDERED','REQUESTED','SHIPPED','INVOICED','SHIPPED VIA & CARRIER CODE'],
                 ['LMR-26-0502','10/26/2026','10/26/2026','10/28/2026','10/28/2026','UPS 2nd Day Air']],
                [['NET SALES AMOUNT','%','%','FREIGHT','SALES TAX','PROCESSING FEE','TOTAL'],
                 ['$252.35','','','$89.5700','','','$341.92']]]

    def preview(self,text,tables=None):
        pdf=MagicMock();page=MagicMock();page.extract_tables.return_value=tables or []
        pdf.__enter__.return_value.pages=[page]
        with patch.object(app,'extract_pdf_text',return_value=text),patch.object(app.pdfplumber,'open',return_value=pdf), \
             patch.object(app,'one',side_effect=AssertionError('No accounting queries')), \
             patch.object(app,'execute',side_effect=AssertionError('No accounting writes')):
            return app.import_vendor_invoice_pdf('invoice.pdf',0,preview_details=True)

    def test_powermation_prices_and_named_freight_cell(self):
        r=self.preview(self.pm,self.tables())
        self.assertEqual((r['invoice_number'],r['date'],r['order_reference']),('30658853','2026-10-28','LMR-26-0502'))
        self.assertEqual([x['item'] for x in r['items']],['TA200DU135','Freight'])
        self.assertEqual([x['amount'] for x in r['items']],[252.35,89.57])
        self.assertEqual(r['parse_status'],'Reconciled lines')
        self.assertAlmostEqual(r['total_due'],341.92)

    def test_powermation_requires_invoiced_date_and_matching_subtotal(self):
        t=self.tables();t[0][1][4]='Unreadable'
        with self.assertRaises(RuntimeError):self.preview(self.pm,t)
        with self.assertRaises(RuntimeError):self.preview(self.pm.replace('1.00 1.00 0.00','1.00 2.00 0.00'),self.tables())
        t=self.tables();t[1][1][0]='$352.35'
        with self.assertRaises(RuntimeError):self.preview(self.pm,t)

    def test_powermation_does_not_guess_which_summary_column_has_a_charge(self):
        t=self.tables();t[1][1][3]='';t[1][1][4]='$89.5700'
        r=self.preview(self.pm,t)
        self.assertEqual(r['items'][-1]['item'],'Sales Tax')
        t=self.tables();t[1][1][3]=''
        with self.assertRaises(RuntimeError):self.preview(self.pm,t)

    def test_border_states_compact_row_quantity_uom_and_total(self):
        r=self.preview(self.bs)
        self.assertEqual((r['invoice_number'],r['date']),('933407913','2026-10-28'))
        self.assertEqual(len(r['items']),1)
        self.assertEqual((r['items'][0]['item'],r['items'][0]['qty'],r['items'][0]['unit']),('126296',1,'EA'))
        self.assertIn('RE108',r['items'][0]['description'])
        self.assertNotIn('Cash discount',r['items'][0]['description'])
        self.assertEqual(r['line_total'],66.74)
        self.assertEqual(r['parse_status'],'Reconciled lines')

    def test_border_states_batch_layout_is_preserved(self):
        text=self.bs.replace('000030 126296 1 EA 1 66.74 /1 EA 66.74',
            '000030 126296 1').replace('506 STOCK','Batch Total: 1 EA 1 66.74 /1 EA 66.74')
        r=self.preview(text)
        self.assertEqual(len(r['items']),1)
        self.assertEqual(r['parse_status'],'Reconciled lines')

    def test_border_states_price_per_hundred_and_separate_freight(self):
        text=self.bs.replace('1 EA 1 66.74 /1 EA 66.74','100 FT 100 66.74 /100 FT 66.74')
        text=text.replace('Handling $ 0.00','Handling $ 10.00').replace('Net Invoice Amount $ 66.74','Net Invoice Amount $ 76.74')
        r=self.preview(text)
        self.assertEqual(r['items'][0]['qty'],100)
        self.assertEqual(r['items'][0]['unit'],'FT')
        self.assertAlmostEqual(r['items'][0]['unit_cost'],.6674)
        self.assertEqual(r['items'][-1]['item'],'Shipping')
        self.assertEqual(r['parse_status'],'Reconciled lines')


if __name__=='__main__':unittest.main()
