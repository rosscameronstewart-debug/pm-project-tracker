import unittest
import gc
import tempfile
from contextlib import closing
from pathlib import Path
from unittest.mock import patch
import app


class SalesForecastTests(unittest.TestCase):
    def test_requested_budget_remaining_formula(self):
        r=app.nte_completion_estimate(97720.47,411234,75)
        self.assertEqual(r['estimated_remaining_sales'],102808.5)
        self.assertEqual(r['estimated_final_sales'],200528.97)

    def test_known_sales_without_guessing_missing_labor(self):
        r=app.nte_completion_estimate(80906.76,411234,75)
        self.assertAlmostEqual(r['estimated_final_sales'],183715.26)

    def test_zero_completion_all_budget_remains(self):
        self.assertEqual(app.nte_completion_estimate(0,1000,0)['estimated_final_sales'],1000)

    def test_finished_work_forecast_equals_sales(self):
        self.assertEqual(app.nte_completion_estimate(200,1000,100)['estimated_final_sales'],200)

    def test_existing_full_budget_override_preserved(self):
        r=app.nte_completion_estimate(200,1000,100,True)
        self.assertEqual(r['estimated_final_sales'],1000)
        self.assertEqual(r['estimated_remaining_sales'],800)
        self.assertEqual(r['estimated_completion_status'],'Full budget override')

    def test_override_only_applies_at_100(self):
        self.assertEqual(app.nte_completion_estimate(200,1000,50,True)['estimated_final_sales'],700)

    def test_over_budget_sales_not_hidden(self):
        r=app.nte_completion_estimate(1200,1000,75)
        self.assertEqual(r['estimated_final_sales'],1450)
        self.assertEqual(r['estimated_completion_variance'],-450)

    def test_sales_basis_for_reported_usage(self):
        self.assertEqual(app.nte_reported_budget_used({'actual_cost':700,'sales_to_date':900,'completion':{}}),900)

    def test_saved_sales_aggregate_excludes_missing_values_and_reconciles_subbuckets(self):
        with tempfile.TemporaryDirectory() as folder:
            root=Path(folder)
            with patch.object(app,'DATA_DIR',root), patch.object(app,'DB_PATH',root/'test.sqlite3'), \
                 patch.object(app,'UPLOAD_DIR',root/'uploads'), patch.object(app,'BRAND_DIR',root/'brand'), \
                 patch.object(app,'seed_bid_tracker_from_workbook'):
                app.init_db()
                with closing(app.db()) as con, con:
                    con.execute("INSERT INTO projects(id,project_code,name,created_at) VALUES (901,'FORECAST-TEST','Forecast test','2026-10-29')")
                    con.execute("INSERT INTO nte_buckets(id,project_id,name,created_at) VALUES (701,901,'Forecast test','2026-10-29')")
                    con.execute("INSERT INTO nte_subbuckets(id,bucket_id,name,cost_type,original_budget,created_at) VALUES (702,701,'Labor','Labor',600,'2026-10-29'),(703,701,'Material','Material',400,'2026-10-29')")
                    con.execute("INSERT INTO cost_records(project_id,nte_subbucket_id,source,cost_type,amount,sales_amount,created_at) VALUES (901,702,'Field Wise','Labor',100,150,'2026-10-29'),(901,702,'Field Wise','Labor',75,0,'2026-10-29'),(901,703,'Vendor Invoice','Material',80,120,'2026-10-29'),(901,703,'Field Wise','Material',-10,-15,'2026-10-29')")
                    con.execute("INSERT INTO nte_completion_updates(project_id,bucket_id,completion_percent,report_date,created_at) VALUES (901,701,75,'2026-10-29','2026-10-29')")
                bucket=app.nte_summary(901,all_targets=True)['buckets'][0]
                self.assertEqual(bucket['sales_to_date'],255)
                self.assertEqual(bucket['missing_sales_count'],1)
                self.assertEqual(bucket['estimated_final_sales'],505)
                self.assertEqual(bucket['estimated_remaining_sales'],250)
                self.assertIn('Provisional',bucket['estimated_completion_status'])
                self.assertEqual(sum(s['estimated_final_sales'] for s in bucket['subbuckets']),505)
                with closing(app.db()) as con:
                    self.assertEqual(con.execute('SELECT SUM(amount) FROM cost_records WHERE project_id=901').fetchone()[0],245)
                metric=app.nte_snapshot_rows({'buckets':[bucket]})[0]
                self.assertEqual(metric['metric_key'],'estimated_final_sales_v2_total')
                self.assertEqual(metric['value'],505)
                gc.collect()

if __name__=='__main__': unittest.main()
