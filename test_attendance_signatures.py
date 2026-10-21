"""Isolated adapter tests: temporary ledger and mocked provider, no real accounts or mail."""
import copy
import tempfile
import unittest
from pathlib import Path
from unittest.mock import MagicMock, Mock, patch
from fastapi import FastAPI, HTTPException
from fastapi.testclient import TestClient
from backend.attendance_signatures import register_signature_preflight
from backend.signature_attempts import SignatureAttempts


class AttendanceSignatureTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.actor = {'username': 'actor', 'role_tags': ['Admin']}
        self.config = {'employes_data': [{'id': 'sheet', 'nom': 'Unit fixture',
            'email_employe': 'fixture@example.com', 'ddc': '2026-01-01',
            'fdc': '2026-12-31', 'responsable': 'Unit manager'}]}
        self.ready = {'can_send': True, 'reason': 'mock sandbox'}
        self.sender = Mock(return_value={'document_id': 'mock-id', 'message': 'Employee then manager'})
        self.policy = Mock()
        self.resolver = patch('backend.attendance_signatures.resolve_manager', return_value={'name': 'Unit manager', 'email': 'manager@example.com'}).start()
        self.addCleanup(patch.stopall)
        self.generator = patch('backend.attendance_signatures.generate_fiche_file', return_value={'data': b'original', 'filename': 'fixture.xlsx'}).start()
        app = FastAPI()
        self.connection = MagicMock()
        self.connection.cursor.return_value.__enter__.return_value.fetchall.return_value = [{'username': 'fixture'}]
        register_signature_preflight(app, get_current_user=lambda: self.actor,
            get_db_connection=Mock(return_value=self.connection), load_user=Mock(), require_target_access=Mock(),
            get_config=lambda *_: copy.deepcopy(self.config), readiness=lambda: self.ready,
            sender=self.sender, attempts=SignatureAttempts(Path(self.temp.name) / 'attempts.sqlite3'), recipient_policy=self.policy)
        self.client = TestClient(app)

    def preview(self, ids=None):
        return self.client.post('/attendance/signature-preview', json={'ids': ids or ['sheet'], 'month': 9, 'year': 2026})

    def payload(self):
        item = self.preview().json()['items'][0]
        return {'id': 'sheet', 'month': 9, 'year': 2026, 'fingerprint': item['fingerprint'], 'request_id': 'unit-attempt-0001'}

    def send(self, payload):
        return self.client.post('/attendance/signature-send', json=payload)

    def test_preflight_is_read_only_and_resolves_manager(self):
        result = self.preview().json()
        self.assertEqual(result['items'][0]['manager_email'], 'manager@example.com')
        self.sender.assert_not_called()
        self.generator.assert_not_called()

    def test_send_reuses_old_service_and_original_file(self):
        result = self.send(self.payload())
        self.assertEqual(result.status_code, 200)
        payload, actor = self.sender.call_args.args
        self.assertEqual(payload.filename, 'fixture.xlsx')
        self.assertEqual(payload.file_b64, 'b3JpZ2luYWw=')
        self.assertEqual(actor, self.actor)
        self.assertEqual(result.json()['status'], 'accepted')

    def test_same_attempt_returns_receipt_without_resend(self):
        payload = self.payload()
        self.assertEqual(self.send(payload).json(), self.send(payload).json())
        self.assertEqual(self.sender.call_count, 1)

    def test_provider_failure_does_not_allow_automatic_retry(self):
        self.sender.side_effect = HTTPException(502, 'mock timeout')
        payload = self.payload()
        self.assertEqual(self.send(payload).status_code, 502)
        self.assertEqual(self.send(payload).status_code, 409)
        self.assertEqual(self.sender.call_count, 1)

    def test_changed_sheet_is_rejected(self):
        payload = self.payload()
        self.config['employes_data'][0]['nom'] = 'Changed unit fixture'
        self.assertEqual(self.send(payload).status_code, 409)
        self.sender.assert_not_called()

    def test_unapproved_live_recipient_is_rejected(self):
        payload = self.payload()
        self.policy.side_effect = HTTPException(403, 'not approved')
        self.assertTrue(self.preview().json()['items'][0]['error'])
        self.assertEqual(self.send(payload).status_code, 403)
        self.sender.assert_not_called()

    def test_mode_change_requires_fresh_confirmation(self):
        payload = self.payload()
        self.ready['mode'] = 'production-test'
        self.assertEqual(self.send(payload).status_code, 409)
        self.sender.assert_not_called()

    def test_changed_manager_is_rejected(self):
        payload = self.payload()
        self.resolver.return_value = {'name': 'Changed manager', 'email': 'changed@example.test'}
        self.assertEqual(self.send(payload).status_code, 409)
        self.sender.assert_not_called()

    def test_missing_configuration_blocks_send(self):
        payload = self.payload()
        self.ready.update(can_send=False, reason='Missing public test URL')
        self.assertEqual(self.send(payload).status_code, 503)
        self.sender.assert_not_called()

    def test_employee_role_cannot_prepare_or_send(self):
        payload = self.payload()
        self.actor['role_tags'] = ['Employe']
        self.assertEqual(self.preview().status_code, 403)
        self.assertEqual(self.send(payload).status_code, 403)

    def test_outside_team_is_rejected_before_provider(self):
        payload = self.payload()
        self.resolver.side_effect = HTTPException(403, 'outside team')
        self.assertEqual(self.send(payload).status_code, 403)
        self.sender.assert_not_called()

    def test_contract_sheet_from_other_month_is_rejected(self):
        self.config['employes_data'][0].update(contract_id=1, sheet_month=10, sheet_year=2026)
        self.assertEqual(self.preview().status_code, 400)

    def test_invalid_row_is_reported_individually(self):
        self.config['employes_data'][0]['email_employe'] = ''
        self.assertTrue(self.preview().json()['items'][0]['error'])

    def test_duplicate_selection_is_deduplicated(self):
        self.assertEqual(len(self.preview(['sheet', 'sheet']).json()['items']), 1)


if __name__ == '__main__':
    unittest.main()
