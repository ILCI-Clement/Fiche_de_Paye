import tempfile
import unittest
from contextlib import nullcontext
from types import SimpleNamespace
from unittest.mock import Mock, patch

from fastapi import FastAPI
from fastapi.testclient import TestClient

from backend.system_storage import register_storage_routes, storage_status


class StorageTests(unittest.TestCase):
    def test_capacity_thresholds_and_shared_disk(self):
        gib = 1024 ** 3
        with tempfile.TemporaryDirectory() as folder:
            for used, free, expected in [(79, 21, 'normal'), (80, 20, 'warning'),
                                         (90, 10, 'critical'), (50, 4, 'critical')]:
                with self.subTest(used=used, free=free), patch('backend.system_storage.shutil.disk_usage',
                        return_value=SimpleNamespace(total=(used + free) * gib, used=used * gib, free=free * gib)) as measure:
                    result = storage_status([('App', folder), ('Documents', folder)])
                    self.assertEqual(len(result['volumes']), 1)
                    self.assertEqual(result['volumes'][0]['level'], expected)
                    self.assertEqual(result['volumes'][0]['labels'], ['App', 'Documents'])
                    self.assertEqual(measure.call_count, 1)

    def test_failed_measurement_has_no_path_or_exception_details(self):
        with patch('backend.system_storage.shutil.disk_usage', side_effect=OSError('private details')):
            result = storage_status([('App', '.')])
        self.assertEqual(result['unavailable'], ['App'])
        self.assertEqual(result['volumes'], [])
        self.assertNotIn('private', str(result))

    def test_authorization_uses_current_roles_and_session(self):
        actor = {'username': 'test', 'session_version': 1, 'role_tags': ['SuperAdmin']}
        current = dict(actor)
        db = Mock()
        db.cursor.return_value = nullcontext(Mock())
        app = FastAPI()
        register_storage_routes(app, get_db_connection=lambda: db, get_current_user=lambda: actor,
                                load_user=lambda *_: current, paths=[('App', '.')])
        with TestClient(app) as client, patch('backend.system_storage.storage_status', return_value={'volumes': []}) as probe:
            for roles, status in [(['Admin'], 200), (['SuperAdmin'], 200), (['Responsable'], 403), (['Employe'], 403)]:
                current['role_tags'] = roles
                response = client.get('/system/storage')
                self.assertEqual(response.status_code, status)
                if status == 200:
                    self.assertEqual(response.headers['cache-control'], 'no-store')
            current['role_tags'] = ['Admin']
            current['session_version'] = 2
            self.assertEqual(client.get('/system/storage').status_code, 401)
            self.assertEqual(probe.call_count, 2)
        self.assertEqual(db.close.call_count, 5)


if __name__ == '__main__':
    unittest.main()
