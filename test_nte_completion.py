import unittest
from io import BytesIO
from unittest.mock import patch
import app


class NteCompletionTests(unittest.TestCase):
    def test_used_totals_mix_forced_and_normal_sections_and_restore_on_clear(self):
        forced = {'current_budget':1000,'actual_cost':100,'sales_to_date':200,'completion':{'completion_percent':100,'use_full_budget':True}}
        normal = {'current_budget':500,'actual_cost':75,'sales_to_date':150,'completion':{'completion_percent':50}}
        self.assertEqual(sum(app.nte_reported_budget_used(b) for b in [forced,normal]),1150)
        self.assertEqual(1500-sum(app.nte_reported_budget_used(b) for b in [forced,normal]),350)
        self.assertEqual(forced['actual_cost'],100)
        forced['completion']['use_full_budget']=False
        self.assertEqual(app.nte_reported_budget_used(forced),200)
        forced['completion'].update(use_full_budget=True,completion_percent=90)
        self.assertEqual(app.nte_reported_budget_used(forced),200)

    def test_reports_show_each_subbucket_completion_and_full_budget_override(self):
        bucket = {'name':'606 - E-Rack - Palermo 13-1', 'current_budget':123912,
                  'actual_cost':20000, 'sales_to_date':28027.28, 'remaining':95884.72,
                  'completion':{'completion_percent':100,'use_full_budget':1}, 'subbuckets':[]}
        bucket.update(app.nte_completion_estimate(28027.28,123912,100,True))
        for name,budget,actual in [('Labor',51912,10793.16),('Material',72000,16934.12)]:
            sub={'name':name,'cost_type':name,'current_budget':budget,'actual_cost':actual/2,'sales_to_date':actual,
                 'remaining':budget-actual,'completion_percent':100,'use_full_budget':True}
            sub.update(app.nte_completion_estimate(actual,budget,100,True))
            bucket['subbuckets'].append(sub)
        data={'buckets':[bucket],'additions':[],'reallocations':[],
              'panel_deliverables':[{'name':'Panel test','bucket_name':bucket['name'],
                                     'total_panels':4,'completed_panels':2,'remaining_panels':2,'completion_percent':50}]}
        with patch.object(app,'one',return_value={'project_code':'TEST','name':'Completion review','customer':'Test customer'}), \
             patch.object(app,'nte_summary',return_value=data), \
             patch.object(app,'save_nte_weekly_kpi_snapshot'), patch.object(app,'nte_kpi_history',return_value=[]):
            html=app.nte_weekly_report_html(1)
            self.assertEqual(html.count('100.0% complete'),5)
            self.assertEqual(html.count('<span style="width:100%">'),3)
            self.assertIn('20.8% of NTE recorded',html)
            self.assertIn('23.5% of NTE recorded',html)
            self.assertIn('Budget override above recorded sales',html)
            self.assertIn('<div class="label">Used</div><div class="value">$123,912.00</div>',html)
            self.assertIn('<div class="label">Remaining</div><div class="value good">$0.00</div>',html)
            self.assertIn('Includes full budget overrides',html)
            self.assertIn('50.0% complete',html)
            self.assertIn('Full budget override',html)
            import pdfplumber
            pdf=app.nte_weekly_report_pdf_bytes(1)
            with pdfplumber.open(BytesIO(pdf)) as report:
                text='\n'.join(p.extract_text() or '' for p in report.pages)
            self.assertEqual(text.count('100.0% complete'),2)
            self.assertIn('Full budget override',text)
            self.assertIn('$123,912.00',text)
            self.assertIn('$123,912.00 $123,912.00 $0.00',text)
            self.report_pdf = pdf

    def test_full_budget_override_matches_screenshot_and_preserves_actual(self):
        actual, budget = 28027.28, 123912
        result = app.nte_completion_estimate(actual, budget, 100, True)
        self.assertEqual(result['estimated_cost_at_completion'], budget)
        self.assertAlmostEqual(result['estimated_cost_to_complete'], 95884.72)
        self.assertEqual(result['estimated_completion_variance'], 0)
        self.assertEqual(result['estimated_completion_status'], 'Full budget override')
        self.assertEqual(actual, 28027.28)

    def test_clearing_override_restores_sales_plus_remaining_budget(self):
        self.assertEqual(app.nte_completion_estimate(200, 1000, 100, False)['estimated_cost_at_completion'], 200)
        self.assertEqual(app.nte_completion_estimate(200, 1000, 50, False)['estimated_cost_at_completion'], 700)
        self.assertEqual(app.nte_completion_estimate(200, 1000, 0, False)['estimated_cost_at_completion'], 1000)

    def test_html_and_pdf_include_unstarted_budget_in_total_and_subbucket(self):
        sub={'name':'Labor','cost_type':'Labor','current_budget':342343,'sales_to_date':0,'completion_percent':0}
        sub.update(app.nte_completion_estimate(0,342343,0))
        bucket={'name':'AC','current_budget':342343,'sales_to_date':0,'completion':{},'subbuckets':[sub]}
        bucket.update(app.nte_completion_estimate(0,342343,0))
        with patch.object(app,'one',return_value={'project_code':'TEST','name':'Zero completion','customer':'Test'}), \
             patch.object(app,'nte_summary',return_value={'buckets':[bucket]}), \
             patch.object(app,'save_nte_weekly_kpi_snapshot'), patch.object(app,'nte_kpi_history',return_value=[]):
            html=app.nte_weekly_report_html(1)
            self.assertIn('<div class="label">Estimated Final Sales</div><div class="value ">$342,343.00</div>',html)
            self.assertIn('Full budget assumed',html)
            import pdfplumber
            with pdfplumber.open(BytesIO(app.nte_weekly_report_pdf_bytes(1))) as report:
                text='\n'.join(p.extract_text() or '' for p in report.pages)
            self.assertIn('$342,343.00',text)
            self.assertIn('Full budget assumed',text)
            self.assertNotIn('Needs field completion percent',text)

    def test_override_tracks_budget_changes_and_subbucket_totals(self):
        labor = app.nte_completion_estimate(10793.16, 51912, 100, True)
        material = app.nte_completion_estimate(16934.12, 72000, 100, True)
        self.assertEqual(labor['estimated_cost_at_completion'] + material['estimated_cost_at_completion'], 123912)
        self.assertEqual(app.nte_completion_estimate(200, 1500, 100, True)['estimated_cost_at_completion'], 1500)

    def test_override_does_not_apply_to_partial_completion_or_negative_remaining(self):
        self.assertEqual(app.nte_completion_estimate(200, 1000, 50, True)['estimated_cost_at_completion'], 700)
        result = app.nte_completion_estimate(1200, 1000, 100, True)
        self.assertEqual(result['estimated_cost_at_completion'], 1000)
        self.assertEqual(result['estimated_cost_to_complete'], 0)


if __name__ == '__main__':
    unittest.main()
