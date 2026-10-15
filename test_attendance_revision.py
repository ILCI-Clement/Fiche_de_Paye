"""Optimistic save tests use an in-memory mock store, not a running database."""
import copy
import json
import unittest
from unittest.mock import MagicMock, patch
from fastapi import HTTPException
from backend.attendance_revision import config_revision, require_revision
import test_backend_helpers
from backend import main


class AttendanceRevisionTests(unittest.TestCase):
    def setUp(self):
        self.state = {'mois': 9, 'annee': 2026, 'employes_data': [], 'unknown': 'keep'}
        self.db = MagicMock()
        self.cursor = self.db.cursor.return_value.__enter__.return_value
        self.cursor.fetchone.side_effect = lambda: {'form_content': json.dumps(self.state)}
        def execute(sql, params=None):
            if 'INSERT INTO Presence' in sql:
                self.state = json.loads(params[1])
        self.cursor.execute.side_effect = execute
        self.link = MagicMock()
        self.sync = MagicMock()
        self.patches = [patch.object(main, 'get_db_connection', return_value=self.db),
                        patch.object(main, 'require_target_access'), patch.object(main, 'validate_manager_save'),
                        patch.object(main, 'link_fiche_records_to_users', self.link),
                        patch.object(main, 'sync_contract_dates_from_config', self.sync)]
        for item in self.patches: item.start()
        self.addCleanup(patch.stopall)
        self.actor = {'username': 'unit', 'role_tags': ['Admin']}

    def test_revision_ignores_dictionary_order(self):
        self.assertEqual(config_revision({'a': 1, 'b': 2}), config_revision({'b': 2, 'a': 1}))

    def test_nested_change_is_detected(self):
        with self.assertRaises(HTTPException):
            require_revision({'nested': {'value': 2}}, config_revision({'nested': {'value': 1}}))

    def test_first_save_returns_exact_written_content_and_revision(self):
        result = main.save_config_checked('unit', copy.deepcopy(self.state), self.actor, config_revision(self.state))
        self.assertEqual(result['config'], self.state)
        self.assertEqual(result['revision'], config_revision(self.state))
        self.db.commit.assert_called_once()

    def test_stale_save_does_not_sync_contracts_or_overwrite(self):
        old_revision = config_revision(self.state)
        self.state['mois'] = 10
        with self.assertRaises(HTTPException) as error:
            main.save_config_checked('unit', {'mois': 8}, self.actor, old_revision)
        self.assertEqual(error.exception.status_code, 409)
        self.assertEqual(self.state['mois'], 10)
        self.link.assert_not_called(); self.sync.assert_not_called()
        self.db.commit.assert_not_called(); self.db.rollback.assert_called_once()

    def test_two_windows_second_save_is_rejected(self):
        original = config_revision(self.state)
        first = {**self.state, 'mois': 10}
        main.save_config_checked('unit', first, self.actor, original)
        with self.assertRaises(HTTPException):
            main.save_config_checked('unit', {**self.state, 'mois': 11}, self.actor, original)
        self.assertEqual(self.state['mois'], 10)

    def test_legacy_save_route_stays_compatible(self):
        self.assertEqual(main.save_config('unit', copy.deepcopy(self.state), self.actor), {'status': 'success'})

    def test_missing_presence_row_uses_empty_revision(self):
        self.cursor.fetchone.side_effect = lambda: None
        result = main.save_config_checked('unit', {'employes_data': []}, self.actor, config_revision({}))
        self.assertEqual(result['revision'], config_revision({'employes_data': []}))

    def test_owner_is_locked_before_presence(self):
        main.save_config_checked('unit', copy.deepcopy(self.state), self.actor, config_revision(self.state))
        statements = [call.args[0] for call in self.cursor.execute.call_args_list]
        self.assertIn('FROM users', statements[0]); self.assertIn('FOR UPDATE', statements[0])
        self.assertIn('FROM Presence', statements[1]); self.assertIn('FOR UPDATE', statements[1])


if __name__ == '__main__':
    unittest.main()
