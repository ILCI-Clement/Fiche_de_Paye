"""PDF and Excel export boundaries on disposable synthetic personnel records."""
from io import BytesIO
from unittest.mock import patch
import json
import unittest

from pypdf import PdfReader
import test_personnel_register_sensitive as fixture


class PersonnelPDFTests(unittest.TestCase):
    setUp = fixture.PersonnelSensitiveHTTPTests.setUp
    create = fixture.PersonnelSensitiveHTTPTests.create
    rows = fixture.PersonnelSensitiveHTTPTests.rows
    raw_update = fixture.PersonnelSensitiveHTTPTests.raw_update

    def export(self, format='pdf', **values):
        return self.client.post(f'/personnel-register/export.{format}', json={'confirm_sensitive_export': True, **values})

    def test_admin_pdf_includes_filtered_all_pages_and_audits_without_plaintext(self):
        for index in range(22):
            self.assertEqual(self.create(full_name=f'PDF-MATCH {index}', social_security_number=fixture.SYNTHETIC_NIR).status_code, 201)
        self.create(full_name='Unrelated fixture')
        response = self.export(q='PDF-MATCH')
        self.assertEqual(response.status_code, 200, response.text[:100] if response.status_code != 200 else '')
        self.assertEqual(response.headers['content-type'], 'application/pdf')
        self.assertIn('no-store', response.headers['cache-control'])
        pdf = PdfReader(BytesIO(response.content))
        text = '\n'.join(page.extract_text() for page in pdf.pages)
        self.assertIn('PDF-MATCH 0', text)
        self.assertIn('PDF-MATCH 21', text)
        self.assertNotIn('Unrelated fixture', text)
        self.assertIn(fixture.SYNTHETIC_NIR, text)
        self.assertFalse(pdf.get_fields())
        events = [row for row in self.rows('personnel_register_events') if row['action'] == 'pdf_exported']
        self.assertEqual(len(events), 22)
        self.assertNotIn(fixture.SYNTHETIC_NIR, json.dumps(events))

    def test_live_super_admin_required_for_excel_and_all_roles_checked_for_pdf(self):
        self.create()
        self.actor['role_tags'] = ['SuperAdmin', 'Admin']  # Stale/forged claim cannot override the DB.
        self.assertEqual(self.export('xlsx').status_code, 403)
        self.raw_update('UPDATE users SET role_tags=? WHERE username=?', ('["SuperAdmin","Admin"]', 'fixture-admin'))
        self.assertEqual(self.export('xlsx').status_code, 200)
        self.assertEqual(self.export().status_code, 200)
        for target in ('fixture-manager', 'fixture-employee'):
            self.actor = {'username': target, 'role_tags': ['SuperAdmin'], 'session_version': 0}
            self.assertEqual(self.export().status_code, 403)
            self.assertEqual(self.export('xlsx').status_code, 403)

    def test_pdf_confirmation_and_audit_failure_do_not_release_data(self):
        self.create(social_security_number=fixture.SYNTHETIC_NIR)
        self.assertEqual(self.export(confirm_sensitive_export=False).status_code, 400)
        self.assertEqual(self.export(q='no-match').status_code, 404)
        self.fail_audit = True
        with self.assertLogs('backend.personnel_register', level='ERROR'):
            result = self.export()
        self.assertEqual(result.status_code, 500)
        self.assertNotIn('attachment', result.headers.get('content-disposition', ''))
        self.assertEqual([row for row in self.rows('personnel_register_events') if row['action'] == 'pdf_exported'], [])

    def test_pdf_generation_failure_returns_no_partial_document(self):
        self.create()
        with patch('backend.personnel_register.personnel_excel_to_pdf', side_effect=ValueError('Synthetic failure')), self.assertLogs('backend.personnel_register', level='ERROR'):
            result = self.export()
        self.assertEqual(result.status_code, 500)
        self.assertNotIn('attachment', result.headers.get('content-disposition', ''))
