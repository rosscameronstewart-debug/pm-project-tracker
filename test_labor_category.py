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


if __name__=='__main__':unittest.main()
