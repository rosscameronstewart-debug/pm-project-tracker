import unittest
from io import BytesIO
from unittest.mock import patch

import pdfplumber
import app


def approval_record():
    return dict(project_code='Palermo 13-1',project_name='Palermo 13-1 Build',customer='Hunt Oil',
                bucket_name='605 - AC - Palermon 13-1 AC',subbucket_name=None,
                from_bucket_name='604 - FC',to_bucket_name='605 - AC',
                status='Pending',amount=10000,labor_amount=0,material_amount=10000,equipment_amount=0,
                requested_date='2026-10-29',approved_date='',approver='Joel Smelly',
                reason='Copper Pricing Increase',
                notes='Written up by Ross. Per Charlie a conversation between Charlie and Shane Swearingen, Hunt has agreed to cover 10k of cable increase costs and the remainder on Twin Peaks.')


class ApprovalSignatureTests(unittest.TestCase):
    def test_both_approval_types_have_two_complete_blank_signature_blocks(self):
        for kind in ('addition','reallocation'):
            with self.subTest(kind=kind), patch.object(app,'one',return_value=approval_record()):
                pdf=app.nte_budget_approval_pdf_bytes(kind,2)
                with pdfplumber.open(BytesIO(pdf)) as doc:
                    self.assertEqual(len(doc.pages),1)
                    text=doc.pages[0].extract_text()
                    self.assertIn('Customer Approval',text)
                    self.assertIn('Twin Peaks Approval',text)
                    self.assertEqual(text.count('Printed Name'),2)
                    self.assertEqual(text.count('Title'),2)
                    # The narrow label wraps onto two lines.
                    self.assertEqual(text.count('Authorized'),2)
                    self.assertEqual(text.count('Signature'),2)
                    self.assertIn('$10,000.00',text)

    def test_long_notes_keep_both_signature_sections_together(self):
        record=approval_record()
        record['notes']='Scope review and supporting explanation. '*180
        with patch.object(app,'one',return_value=record):
            with pdfplumber.open(BytesIO(app.nte_budget_approval_pdf_bytes('addition',2))) as doc:
                texts=[page.extract_text() or '' for page in doc.pages]
        customer=[i for i,text in enumerate(texts) if 'Customer Approval' in text]
        contractor=[i for i,text in enumerate(texts) if 'Twin Peaks Approval' in text]
        self.assertEqual(len(customer),1)
        self.assertEqual(customer,contractor)
        self.assertIn('Printed Name',texts[customer[0]])
        self.assertIn('Please sign and return',texts[customer[0]])


if __name__=='__main__':
    unittest.main()
