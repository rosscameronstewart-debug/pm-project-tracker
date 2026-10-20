import unittest
from io import BytesIO
from unittest.mock import patch
import app


class NteCompletionTests(unittest.TestCase):
    def test_reports_show_each_subbucket_completion_and_full_budget_override(self):
        bucket = {'name':'606 - E-Rack - Palermo 13-1', 'current_budget':123912,
                  'actual_cost':28027.28, 'remaining':95884.72,
                  'completion':{'completion_percent':100,'use_full_budget':1}, 'subbuckets':[]}
        bucket.update(app.nte_completion_estimate(28027.28,123912,100,True))
        for name,budget,actual in [('Labor',51912,10793.16),('Material',72000,16934.12)]:
            sub={'name':name,'cost_type':name,'current_budget':budget,'actual_cost':actual,
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
            self.assertIn('20.8% of budget spent',html)
            self.assertIn('23.5% of budget spent',html)
            self.assertIn('Projected cost not yet recorded',html)
            self.assertIn('50.0% complete',html)
            self.assertIn('Full budget override',html)
            from pypdf import PdfReader
            pdf=app.nte_weekly_report_pdf_bytes(1)
            text='\n'.join(p.extract_text() for p in PdfReader(BytesIO(pdf)).pages)
            self.assertEqual(text.count('100.0% complete'),2)
            self.assertIn('Full budget override',text)
            self.assertIn('$123,912.00',text)
            self.report_pdf = pdf

    def test_full_budget_override_matches_screenshot_and_preserves_actual(self):
        actual, budget = 28027.28, 123912
        result = app.nte_completion_estimate(actual, budget, 100, True)
        self.assertEqual(result['estimated_cost_at_completion'], budget)
        self.assertAlmostEqual(result['estimated_cost_to_complete'], 95884.72)
        self.assertEqual(result['estimated_completion_variance'], 0)
        self.assertEqual(result['estimated_completion_status'], 'Full budget override')
        self.assertEqual(actual, 28027.28)

    def test_clearing_override_restores_actual_based_projection(self):
        self.assertEqual(app.nte_completion_estimate(200, 1000, 100, False)['estimated_cost_at_completion'], 200)
        self.assertEqual(app.nte_completion_estimate(200, 1000, 50, False)['estimated_cost_at_completion'], 400)
        self.assertEqual(app.nte_completion_estimate(200, 1000, 0, False)['estimated_cost_at_completion'], 0)

    def test_override_tracks_budget_changes_and_subbucket_totals(self):
        labor = app.nte_completion_estimate(10793.16, 51912, 100, True)
        material = app.nte_completion_estimate(16934.12, 72000, 100, True)
        self.assertEqual(labor['estimated_cost_at_completion'] + material['estimated_cost_at_completion'], 123912)
        self.assertEqual(app.nte_completion_estimate(200, 1500, 100, True)['estimated_cost_at_completion'], 1500)

    def test_override_does_not_apply_to_partial_completion_or_negative_remaining(self):
        self.assertEqual(app.nte_completion_estimate(200, 1000, 50, True)['estimated_cost_at_completion'], 400)
        result = app.nte_completion_estimate(1200, 1000, 100, True)
        self.assertEqual(result['estimated_cost_at_completion'], 1000)
        self.assertEqual(result['estimated_cost_to_complete'], 0)


if __name__ == '__main__':
    unittest.main()
