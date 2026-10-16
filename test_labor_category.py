import unittest
from unittest.mock import patch
import app


class LaborCategoryTests(unittest.TestCase):
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
        self.assertEqual(app.labor_category(records[0]['item']), 'Automation Tech 1 ST')

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
