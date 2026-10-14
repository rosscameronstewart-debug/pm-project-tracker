import sqlite3,unittest
from datetime import datetime
import app
class PoNumberTests(unittest.TestCase):
    def test_customer_code_resolution(self):
        con=sqlite3.connect(':memory:');con.row_factory=sqlite3.Row
        con.execute("CREATE TABLE cog_categories(id INTEGER,name TEXT,po_code TEXT,active INTEGER,customer_name TEXT DEFAULT '')")
        con.executemany('INSERT INTO cog_categories(id,name,po_code,active) VALUES (?,?,?,1)',[(1,'Chord Maintenance','CRD'),(2,'Kraken Maintenance','KRKN'),(3,'CPX Bakken Maintenance','CPX'),(4,'CPX Texas Maintenance','CPX'),(5,'Blank','')])
        for customer,code in [('Chord Energy','CRD'),('Kraken','KRKN'),('CPX ND','CPX')]:
            self.assertEqual(app.po_code_for_customer(con,{},customer),code)
        self.assertEqual(app.po_code_for_customer(con,{'cog_category_id':2},'Chord'),'KRKN')
        for ref,customer in [({},'Unknown'),({'cog_category_id':5},'Chord')]:
            with self.assertRaises(ValueError):app.po_code_for_customer(con,ref,customer)
    def test_shared_sequence_preserves_existing(self):
        con=sqlite3.connect(':memory:');con.row_factory=sqlite3.Row
        con.execute('CREATE TABLE purchase_orders(po_number TEXT UNIQUE)')
        year=datetime.now().year
        short_year=f'{year % 100:02d}'
        for number in [f'PO-{year}-0381',f'AMR-{year}-0382',f'OLD-{year-1}-9000']:
            con.execute('INSERT INTO purchase_orders VALUES (?)',(number,))
        con.commit()
        self.assertEqual(app.next_po_number(con,'trk'),f'TRK-{short_year}-0383')
        con.execute('INSERT INTO purchase_orders VALUES (?)',(f'TRK-{short_year}-0383',))
        self.assertEqual(app.next_po_number(con,'001'),f'001-{short_year}-0384')
        self.assertEqual(con.execute('SELECT count(*) FROM purchase_orders').fetchone()[0],4)
        con.execute('INSERT INTO purchase_orders VALUES (?)',(f'XYZ-{year}-9999',))
        with self.assertRaises(ValueError):app.next_po_number(con,'A')
if __name__=='__main__':unittest.main()
