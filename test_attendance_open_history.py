"""Opening saved monthly attendance must not recreate or alter contract history."""
import copy
import json
import unittest
from unittest.mock import MagicMock, Mock

from fastapi import FastAPI
from fastapi.testclient import TestClient

from backend.attendance_api import register_attendance_routes
from backend.attendance_revision import config_revision


class AttendanceOpenHistoryTests(unittest.TestCase):
    def setUp(self):
        self.actor = {'username': 'operator', 'role_tags': ['Admin']}
        self.account = {'username': 'opaque-person', 'email': 'person@example.test',
                        'full_name': 'Synthetic Person', 'manager_username': None}
        self.saved = {'id': 'existing-september', 'account_username': 'opaque-person',
                      'contract_id': 7, 'contract_snapshot': {'id': 7, 'start_date': '2025-09-15', 'end_date': '2026-09-14'},
                      'sheet_month': 9, 'sheet_year': 2026, 'sheet_period_start': '2026-09-01',
                      'sheet_period_end': '2026-09-14', 'nom': 'Historical export name',
                      'calendar_overrides': {'2026-09-02': {'matin': ['CP', '']}}}
        self.config = {'mois': 10, 'annee': 2026, 'employes_data': [copy.deepcopy(self.saved)]}
        self.contracts = []
        self.access = Mock(return_value=True)
        self.save = Mock(side_effect=lambda _owner, config, _actor, _revision:
                         {'config': config, 'revision': config_revision(config)})
        self.db = MagicMock()
        cursor = self.db.cursor.return_value.__enter__.return_value
        cursor.fetchone.side_effect = lambda: {'form_content': json.dumps(self.config)}
        cursor.fetchall.side_effect = lambda: copy.deepcopy(self.contracts)
        app = FastAPI()
        register_attendance_routes(app, get_db_connection=lambda: self.db,
            get_current_user=lambda: self.actor, load_user=lambda *_: self.account,
            can_manage=self.access, save_config=Mock(), save_checked=self.save)
        self.client = TestClient(app)
        self.addCleanup(self.client.close)

    def open(self, month=9, revision=None):
        return self.client.post('/attendance/open/opaque-person', json={
            'month': month, 'year': 2026,
            'revision': revision or config_revision(self.config)})

    def test_saved_month_opens_without_current_contract_and_keeps_snapshot(self):
        response = self.open()
        self.assertEqual(response.status_code, 200, response.text)
        self.assertEqual(response.json()['active_id'], self.saved['id'])
        self.assertEqual(response.json()['config']['employes_data'], [self.saved])
        self.assertEqual(response.json()['config']['mois'], 9)
        self.assertEqual(self.config['employes_data'], [self.saved])

    def test_archived_contract_does_not_hide_existing_monthly_snapshot(self):
        self.contracts = [{'id': 7, 'start_date': '2025-09-15', 'end_date': '2026-09-14', 'is_archived': True}]
        self.assertEqual(self.open().status_code, 200)
        self.assertEqual(self.save.call_args.args[1]['employes_data'], [self.saved])

    def test_changed_contract_dates_do_not_rewrite_saved_snapshot(self):
        self.contracts = [{'id': 7, 'start_date': '2027-01-01', 'end_date': None}]
        self.assertEqual(self.open().status_code, 200)
        self.assertEqual(self.save.call_args.args[1]['employes_data'], [self.saved])

    def test_missing_contract_cannot_create_another_month(self):
        self.assertEqual(self.open(month=10).status_code, 409)
        self.save.assert_not_called()

    def test_other_account_history_cannot_be_opened_by_same_display_name(self):
        self.config['employes_data'][0]['account_username'] = 'different-person'
        self.assertEqual(self.open().status_code, 409)
        self.save.assert_not_called()

    def test_archived_sheet_still_requires_restore(self):
        self.config['employes_data'][0]['is_archived'] = True
        self.assertEqual(self.open().status_code, 409)
        self.save.assert_not_called()

    def test_outside_team_still_rejected_before_opening_history(self):
        self.actor['role_tags'] = ['Responsable']
        self.access.return_value = False
        self.assertEqual(self.open().status_code, 403)
        self.save.assert_not_called()

    def test_employee_role_cannot_open_history_in_editor(self):
        self.actor['role_tags'] = ['Employe']
        self.assertEqual(self.open().status_code, 403)
        self.save.assert_not_called()

    def test_stale_revision_cannot_overwrite_workspace(self):
        self.assertEqual(self.open(revision='0' * 64).status_code, 409)
        self.save.assert_not_called()


if __name__ == '__main__':
    unittest.main()
