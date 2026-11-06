"""Verify staged entry wiring without loading credentials, a database, or SMTP."""
import importlib.util
from pathlib import Path
from types import SimpleNamespace
import unittest
from unittest.mock import Mock, patch
import backend


class RebuildServerTests(unittest.TestCase):
    def test_staged_entry_reuses_authorization_and_versioned_save(self):
        names = ['get_db_connection', 'get_current_user', 'load_user', 'user_can_manage',
                 'save_config', 'save_config_checked', 'require_target_access', 'get_config', 'send_fiche_for_signature',
                 'insert_account', 'welcome_created_account']
        services = SimpleNamespace(app=object(), TRANSPORT_RECEIPTS_DIR=Path('/synthetic-documents'),
                                   ESIGN_DOCUMENTS_DIR=Path('/synthetic-signing'),
                                   **{name: Mock(name=name) for name in names})
        spec = importlib.util.spec_from_file_location('backend._staged_entry_test', Path(__file__).parent / 'backend/rebuild_server.py')
        module = importlib.util.module_from_spec(spec)
        with patch.object(backend, 'main', services, create=True), \
                patch('backend.attendance_api.register_attendance_routes') as attendance, \
                patch('backend.attendance_signatures.register_signature_preflight') as signatures, \
                patch('backend.person_profiles.register_profile_routes') as profiles, \
                patch('backend.personnel_register.register_personnel_routes') as personnel, \
                patch('backend.medical_visits.register_medical_routes') as medical, \
                patch('backend.resource_maintenance.register_resource_maintenance') as maintenance, \
                patch('backend.personnel_snapshots.register_snapshot_routes') as snapshots, \
                patch('backend.system_storage.register_storage_routes') as storage, \
                patch('backend.signature_attempts.SignatureAttempts') as attempts:
            spec.loader.exec_module(module)
            self.assertIs(module.app, services.app)
            self.assertIs(storage.call_args.kwargs['get_current_user'], services.get_current_user)
            self.assertIs(attendance.call_args.kwargs['save_checked'], services.save_config_checked)
            self.assertIs(attendance.call_args.kwargs['get_current_user'], services.get_current_user)
            self.assertIs(signatures.call_args.kwargs['require_target_access'], services.require_target_access)
            self.assertIs(signatures.call_args.kwargs['sender'], services.send_fiche_for_signature)
            self.assertIs(signatures.call_args.kwargs['attempts'], attempts.return_value)
            self.assertIs(profiles.call_args.kwargs['get_current_user'], services.get_current_user)
            self.assertIs(personnel.call_args.args[0], services.app)
            self.assertIs(personnel.call_args.kwargs['get_db_connection'], services.get_db_connection)
            self.assertIs(personnel.call_args.kwargs['get_current_user'], services.get_current_user)
            self.assertIs(personnel.call_args.kwargs['load_user'], services.load_user)
            self.assertIs(personnel.call_args.kwargs['insert_account'], services.insert_account)
            self.assertIs(personnel.call_args.kwargs['welcome_account'], services.welcome_created_account)
            self.assertIs(medical.call_args.kwargs['get_current_user'], services.get_current_user)
            self.assertIs(medical.call_args.kwargs['load_user'], services.load_user)
            self.assertIs(medical.call_args.kwargs['can_manage'], services.user_can_manage)
            self.assertIs(snapshots.call_args.kwargs['get_current_user'], services.get_current_user)
            self.assertIs(snapshots.call_args.kwargs['load_user'], services.load_user)
            self.assertIs(maintenance.call_args.kwargs['get_current_user'], services.get_current_user)
            self.assertIs(maintenance.call_args.kwargs['load_user'], services.load_user)
        for name in names:
            getattr(services, name).assert_not_called()


if __name__ == '__main__':
    unittest.main()
