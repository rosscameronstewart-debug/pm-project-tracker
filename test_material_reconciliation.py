import sqlite3
import unittest
import material_reconciliation as rec

def row(ticket='1',day='2026-10-25',qty=2,status='Submitted',job='204'):
    return dict(ticket_number=ticket,ticket_date=day,order_number=job,customer='Example',item='Widget DSG 128',description='Widget',quantity=qty,rate=20,amount=qty*20,status=status)

class RollingComparisonTests(unittest.TestCase):
    def setUp(self):
        self.c=sqlite3.connect(':memory:');self.c.row_factory=sqlite3.Row
        self.addCleanup(self.c.close);rec.init_schema(self.c)
    def ingest(self,rows,end='2026-10-26',sha='one'):
        with self.c:return rec.import_sales_rows(self.c,rows,end,'example.xlsx',sha,'test')
    def test_identical_and_overlapping_uploads_do_not_double_sales(self):
        self.ingest([row()]);self.assertTrue(self.ingest([row()])['duplicate_file'])
        self.ingest([row()],end='2026-10-27',sha='two')
        self.assertEqual(len(rec.active_sales_rows(self.c)),1)
        self.assertEqual(self.c.execute('SELECT COUNT(*) FROM material_sales_ticket_history').fetchone()[0],1)
    def test_real_duplicate_rows_are_preserved(self):
        self.ingest([row(),row()]);sales=rec.active_sales_rows(self.c)
        self.assertEqual(len(sales),2);self.assertNotEqual(sales[0]['sales_line_id'],sales[1]['sales_line_id'])
    def test_revised_quantity_and_job_replace_ticket_and_keep_history(self):
        self.ingest([row()]);self.ingest([row(qty=3,job='205')],sha='two')
        sales=rec.active_sales_rows(self.c)
        self.assertEqual(len(sales),1);self.assertEqual(sales[0]['quantity'],3);self.assertEqual(sales[0]['order_number'],'205')
        self.assertEqual(self.c.execute('SELECT COUNT(*) FROM material_sales_ticket_history').fetchone()[0],2)
    def test_missing_ticket_in_complete_window_inactivated_but_old_history_retained(self):
        self.ingest([row(ticket='old',day='2026-10-20'),row(ticket='gone')])
        self.ingest([],end='2026-10-27',sha='two')
        self.assertEqual([r['ticket_number'] for r in rec.active_sales_rows(self.c)],['OLD'])
        self.assertEqual(self.c.execute('SELECT COUNT(*) FROM material_sales_ticket_history').fetchone()[0],3)
    def test_draft_and_closed_do_not_count_as_sold(self):
        self.ingest([row('1',status='Draft'),row('2',status='Closed'),row('3',status='Approved'),row('4',status='Invoiced'),row('5',status='Submitted')])
        self.assertEqual(sum(s['counts_as_sold'] for s in rec.active_sales_rows(self.c)),3)
    def test_wrong_window_and_stale_export_rejected(self):
        with self.assertRaises(ValueError):self.ingest([row(day='2026-10-01')])
        self.ingest([row()])
        with self.assertRaises(ValueError):self.ingest([row(day='2026-10-19')],end='2026-10-20',sha='older')
    def test_quantity_not_reused_and_revised_sales_invalidate_confirmation(self):
        self.ingest([row(qty=2)])
        sales=rec.active_sales_rows(self.c);sale=sales[0]
        buys=[{'purchase_key':key,'state':'Candidate sale — verify units and quantity','bought_qty':2,'candidates':[sale]} for key in ('buy1','buy2')]
        fields=dict(purchase_key='buy1',sales_line_id=sale['sales_line_id'],bought_qty=2,bought_units_per_sale_unit=1,unit_note='EA to EA verified')
        with self.c: rec.confirm_match(self.c,buys,sales,fields,'test')
        with self.assertRaises(ValueError):rec.confirm_match(self.c,buys,sales,{**fields,'purchase_key':'buy2'},'test')
        self.ingest([row(qty=1)],sha='changed')
        current,stale=rec.valid_matches(self.c,buys,rec.active_sales_rows(self.c))
        self.assertEqual(len(current),0);self.assertEqual(len(stale),1)
    def test_unit_conversion_must_be_explicit(self):
        self.ingest([row()]);sales=rec.active_sales_rows(self.c)
        buy={'purchase_key':'buy','state':'Candidate sale — verify units and quantity','bought_qty':2,'candidates':sales}
        with self.assertRaises(ValueError):rec.confirm_match(self.c,[buy],sales,dict(purchase_key='buy',sales_line_id=sales[0]['sales_line_id'],bought_qty=2,bought_units_per_sale_unit=1),'test')

    def test_transition_between_accepted_sold_statuses_keeps_line_identity(self):
        self.ingest([row(status='Submitted')]);before=rec.active_sales_rows(self.c)[0]['sales_line_id']
        self.ingest([row(status='Invoiced')],sha='changed-status')
        after=rec.active_sales_rows(self.c)[0]
        self.assertEqual(before,after['sales_line_id']);self.assertTrue(after['counts_as_sold'])

    def test_detailed_invoice_preview_exposes_prices_without_importing(self):
        from unittest.mock import patch
        import app
        parsed=dict(vendor='Amazon',invoice_number='123',invoice_date='2026-10-20',order_number='PO-1',total_due=10,item_lines=[dict(product_code='PART',description='Widget',qty=2,unit_price=5,amount=10)])
        with patch.object(app,'extract_pdf_text',return_value='Final Details for Order #123 amazon.com'),patch.object(app,'parse_amazon_order_invoice',return_value=parsed),patch.object(app,'execute',side_effect=AssertionError('Preview must not write')),patch.object(app,'one',side_effect=AssertionError('Preview must not query import targets')):
            result=app.import_vendor_invoice_pdf('example.pdf',1,preview_details=True)
        self.assertEqual(result['items'][0]['unit_cost'],5)
        self.assertEqual(result['items'][0]['amount'],10)
        self.assertEqual(result['parse_status'],'Reconciled lines')

class PurchaseTimingTests(unittest.TestCase):
    def test_forward_only_baseline_and_carry_forward(self):
        self.assertFalse(rec.in_forward_workflow({'invoice_date':'2026-10-26','uploaded_at':'2026-10-28T12:00:00'},'2026-10-27'))
        self.assertTrue(rec.in_forward_workflow({'invoice_date':'2026-10-27'},'2026-10-27'))
        self.assertTrue(rec.in_forward_workflow({'invoice_date':'2026-11-15'},'2026-10-27'))
        # Scope uses a fixed starting date, never a rolling seven-day cutoff.
        self.assertTrue(rec.in_forward_workflow({'invoice_date':'2026-10-27'},'2026-10-27'))
    def test_new_undated_upload_needs_review_but_old_undated_backlog_excluded(self):
        self.assertTrue(rec.in_forward_workflow({'invoice_date':None,'uploaded_at':'2026-10-27T10:00:00'},'2026-10-27'))
        self.assertFalse(rec.in_forward_workflow({'invoice_date':'bad','uploaded_at':'2026-10-26T10:00:00'},'2026-10-27'))
        self.assertFalse(rec.in_forward_workflow({},'2026-10-27'))
    def assess(self,day,coverage=None,**extra):
        from datetime import date
        purchase=dict(invoice_date=day,bought_qty=10,confirmed_qty=0,state='No exact item link in available sales')
        purchase.update(extra)
        if coverage is None:coverage=['2026-10-'+str(d) for d in range(20,27)]
        return rec.purchase_timing(purchase,coverage,date(2026,10,27))
    def test_grace_and_overdue_boundary_uses_loaded_day(self):
        self.assertEqual(self.assess('2026-10-23')['timing_state'],'Within 0–3 days — waiting for sale')
        result=self.assess('2026-10-22')
        self.assertEqual(result['age_days'],4)
        self.assertEqual(result['timing_as_of'],'2026-10-26')
        self.assertEqual(result['timing_state'],'Over 3 days — review unconfirmed quantity')
    def test_missing_or_gapped_history_is_not_overdue(self):
        self.assertEqual(self.assess('2026-10-19')['timing_state'],'Insufficient sales history')
        self.assertEqual(self.assess('2026-10-20',['2026-10-20','2026-10-26'])['timing_state'],'Insufficient sales history')
        self.assertEqual(self.assess('2026-10-20',[])['timing_state'],'Insufficient sales history')
    def test_missing_future_and_confirmed_dates(self):
        self.assertEqual(self.assess(None)['timing_state'],'Cannot determine — invoice date missing')
        self.assertEqual(self.assess('2026-10-27')['timing_state'],'Purchase after loaded sales date')
        self.assertEqual(self.assess('2026-10-20',confirmed_qty=10)['timing_state'],'Confirmed sold')
    def test_partial_quantity_and_source_issues(self):
        self.assertEqual(self.assess('2026-10-20',confirmed_qty=2)['timing_state'],'Over 3 days — review unconfirmed quantity')
        self.assertEqual(self.assess('2026-10-20',state='Needs invoice parsing',bought_qty=None)['timing_state'],'Over 3 days — resolve source issue')

if __name__=='__main__':unittest.main()
