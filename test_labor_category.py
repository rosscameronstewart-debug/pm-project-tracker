import unittest
from unittest.mock import patch, MagicMock
import app


class LaborCategoryTests(unittest.TestCase):
    def test_pdf_import_stores_customer_sales_once_and_keeps_internal_cost_separate(self):
        page = MagicMock()
        page.extract_text.return_value = ('Field Ticket #16346\nJob # 604 Palermo\nTicket Date 10/2/2026\n'
            'Work Description\nInstall wiring\nLabor\nTime Entries\nName Date\n'
            'Jared Brown 10/2/2026 Journeyman Reg 8.00\nBillable Entries\n')
        page.extract_tables.return_value = [[['Work Type / Comp Item','Qty','Rate','Total'],
            ['Jared Brown - Journeyman - Reg','8.00','$97.00','$776.00']]]
        pdf = MagicMock()
        pdf.__enter__.return_value.pages = [page]
        with patch.object(app.pdfplumber, 'open', return_value=pdf), \
             patch.object(app, 'one', return_value=None), patch.object(app, 'execute') as execute, \
             patch.object(app, 'raw_rate_for', return_value={'raw_rate':81.02,'category':'Journeyman ST','source':'Project rate set'}):
            result = app.import_fieldwise_pdf('ticket16346.pdf',1006)
        inserts = [call for call in execute.call_args_list if 'INSERT INTO cost_records' in call.args[0]]
        self.assertEqual(result['count'],1)
        self.assertEqual(len(inserts),1)
        values = inserts[0].args[1]
        self.assertEqual(values[10],8)
        self.assertAlmostEqual(values[11],81.02)
        self.assertAlmostEqual(values[12],648.16)
        self.assertEqual(values[13:15],(97,776))

    def test_billable_labor_takes_precedence_over_duplicate_time_entries(self):
        billed = [dict(cost_type='Labor', item='Jared Brown - Journeyman - Reg',
                       qty=8, rate=97, amount=776)]
        time = [dict(billed[0], rate=0, amount=0)]
        result = app.fieldwise_labor_with_time_fallback(billed, time)
        self.assertEqual(result, billed)
        self.assertEqual(sum(r['amount'] for r in result), 776)

    def test_overtime_name_spacing_does_not_create_second_labor_entry(self):
        billed = [dict(cost_type='Labor', item='Jared Brown - Journeyman OT',
                       qty=1, rate=145.5, amount=145.5)]
        time = [dict(billed[0], item='Jared Brown - Journeyman - OT', rate=0, amount=0)]
        self.assertEqual(app.fieldwise_labor_with_time_fallback(billed, time), billed)

    def test_printed_billable_hours_win_over_conflicting_time_hours(self):
        billed = [dict(cost_type='Labor', item='Jared Brown - Journeyman - Reg',
                       qty=5, rate=97, amount=485)]
        time = [dict(billed[0], qty=4, rate=0, amount=0)]
        self.assertEqual(app.fieldwise_labor_with_time_fallback(billed, time), billed)

    def test_unbilled_people_and_classes_remain_available_for_review(self):
        billed = [dict(cost_type='Labor', item='Jared Brown - Journeyman - Reg',
                       qty=8, rate=97, amount=776)]
        missing = dict(billed[0], item='Other Person - Journeyman - Reg', rate=0, amount=0)
        self.assertEqual(app.fieldwise_labor_with_time_fallback(billed, [missing]), billed + [missing])

    def test_confirmed_aliases(self):
        for item,expected in [
            ('Jerry Nguyen - Panel Shop Labor','Panel Shop ST'),
            ('Landen Silvis - Panel Shop OT Labor','Panel Shop OT'),
            ('Jerry Nguyen - panel shop ot labor','Panel Shop OT'),
            ('Employee - Panel Shop - Reg','Panel Shop ST'),
            ('Nate Spence - Panel Shop 8.00 Apprentice 4 - OT','Panel Shop 8.00 Apprentice 4 OT'),
            ('Nate Spence - Office Staff - - OT','Office Staff OT')]:
            with self.subTest(item=item):self.assertEqual(app.labor_category(item),expected)

    def test_lookup_uses_normalized_category(self):
        with patch.object(app,'one',side_effect=[{'rate_set_id':1},{'raw_rate':50}]) as query:
            result=app.raw_rate_for('Labor','Jerry Nguyen - Panel Shop OT Labor',1001)
            self.assertEqual(result,{'category':'Panel Shop OT','raw_rate':50,'source':'Project rate set'})
            self.assertEqual(query.call_args.args[1],('Labor','Panel Shop OT',1,1))

    def test_unmarked_time_entries_default_to_regular_preserving_class_number(self):
        text = ('Time Entries\nName Date\n'
                'Miguel Mireles 9/21/2026 Automation Tech 1 8.00\n'
                'Miguel Mireles 9/22/2026 Automation Tech 1 4.00\nSUBTOTAL $0.00')
        records = app.fieldwise_time_entry_labor_records_from_text(text, 'Panel work')
        self.assertEqual([r['qty'] for r in records], [8, 4])
        self.assertTrue(all(r['item'] == 'Miguel Mireles - Automation Tech 1 - Reg' for r in records))
        self.assertEqual(app.labor_category(records[0]['item']), 'Automation Tech ST')

    def test_automation_alias_only_maps_confirmed_regular_category(self):
        self.assertEqual(app.labor_category('Miguel Mireles - Automation Tech 1 - Reg'), 'Automation Tech ST')
        self.assertEqual(app.labor_category('Miguel Mireles - Automation Tech 1 - OT'), 'Automation Tech 1 OT')

    def test_explicit_regular_and_overtime_still_take_precedence(self):
        text = ('Time Entries\nName Date\n'
                'Employee 9/21/2026 Panel Shop Reg 8.00 Panel Shop OT 2.00\nSUBTOTAL $0.00')
        records = app.fieldwise_time_entry_labor_records_from_text(text, '')
        self.assertEqual([(r['item'], r['qty']) for r in records],
                         [('Employee - Panel Shop - Reg', 8), ('Employee - Panel Shop - OT', 2)])

    def test_unmarked_integer_hours_and_zero_or_missing_hours(self):
        text = ('Time Entries\nName Date\n'
                'Employee 9/21/2026 Automation Tech 1 8\n'
                'Employee 9/22/2026 Automation Tech 1 0.00\n'
                'Employee 9/23/2026 Panel Shop\nSUBTOTAL $0.00')
        records = app.fieldwise_time_entry_labor_records_from_text(text, '')
        self.assertEqual(len(records), 1)
        self.assertEqual(records[0]['qty'], 8)
        self.assertEqual(records[0]['item'], 'Employee - Automation Tech 1 - Reg')


if __name__=='__main__':unittest.main()
