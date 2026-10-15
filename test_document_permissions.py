"""Documents authorization with temporary files and mocked database connections."""
import base64
from pathlib import Path
import tempfile
import unittest
import io
import zipfile
from unittest.mock import MagicMock, patch

from fastapi import FastAPI, HTTPException
from fastapi.testclient import TestClient
from backend.documents import register_document_routes


class DocumentPermissionTests(unittest.TestCase):
    def test_server_zip_uses_authorized_download_and_deduplicates(self):
        self.actor = {'username':'admin','role_tags':['Admin']}
        self.cursor.fetchone.return_value = self.row('employee')
        (self.storage/'stored.pdf').write_bytes(self.content)
        response=self.client.post('/transport-receipts/export',json={'ids':[11,11]})
        self.assertEqual(response.status_code,200)
        with zipfile.ZipFile(io.BytesIO(response.content)) as archive:
            self.assertEqual(len(archive.namelist()),1)
            self.assertEqual(archive.read(archive.namelist()[0]),self.content)

    def test_server_zip_rejects_unauthorized_file(self):
        self.cursor.fetchone.return_value = self.row('outsider')
        response=self.client.post('/transport-receipts/export',json={'ids':[11]})
        self.assertEqual(response.status_code,403)

    def test_server_zip_missing_file_does_not_return_partial_archive(self):
        self.cursor.fetchone.return_value = self.row('employee')
        response=self.client.post('/transport-receipts/export',json={'ids':[11]})
        self.assertEqual(response.status_code,404)
        self.assertNotEqual(response.headers.get('content-type'),'application/zip')

    def test_server_zip_rejects_empty_and_oversized_selection(self):
        self.assertEqual(self.client.post('/transport-receipts/export',json={'ids':[]}).status_code,400)
        self.assertEqual(self.client.post('/transport-receipts/export',json={'ids':list(range(101))}).status_code,400)

    def test_server_zip_enforces_content_size_limit(self):
        self.cursor.fetchone.return_value = self.row('employee')
        (self.storage/'stored.pdf').write_bytes(self.content)
        with patch('backend.documents.MAX_ARCHIVE_BYTES',1):
            response=self.client.post('/transport-receipts/export',json={'ids':[11]})
        self.assertEqual(response.status_code,400)

    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.storage = Path(self.temp.name)
        self.actor = {'username': 'employee', 'role_tags': ['Employe']}
        self.cursor = MagicMock()
        self.connection = MagicMock()
        self.connection.cursor.return_value.__enter__.return_value = self.cursor
        self.connect = MagicMock(return_value=self.connection)
        self.cursor.lastrowid = 11
        self.load_user = MagicMock(side_effect=lambda cursor, name: {'username': name})
        self.manage = MagicMock(side_effect=lambda cursor, actor, target: target['username'] == 'team-member')

        def authenticated():
            if self.actor is None:
                raise HTTPException(401, 'Authentication required')
            return self.actor

        def require_roles(*roles):
            def allowed():
                actor = authenticated()
                if not set(roles).intersection(actor['role_tags']):
                    raise HTTPException(403, 'Forbidden')
                return actor
            return allowed

        app = FastAPI()
        register_document_routes(
            app, get_db_connection=self.connect, get_current_user=authenticated,
            require_roles=require_roles, load_user=self.load_user, user_can_manage=self.manage,
            all_users=lambda cursor: [{'username': 'team-member'}, {'username': 'outsider'}],
            storage_dir=self.storage, max_file_bytes=10 * 1024 * 1024,
        )
        self.client = TestClient(app)
        self.addCleanup(self.client.close)
        self.content = b'%PDF-1.4\nunit-test\n%%EOF'

    def row(self, owner):
        return dict(id=11, employee_username=owner, uploaded_by=self.actor['username'],
                    original_filename='invoice.pdf', stored_filename='stored.pdf',
                    mime_type='application/pdf', size_bytes=len(self.content),
                    document_type='transport', receipt_month=9, receipt_year=2026,
                    archived=0, created_at='2026-10-14T12:00:00')

    def payload(self, owner):
        return dict(employee_username=owner, original_filename='invoice.pdf',
                    file_b64=base64.b64encode(self.content).decode(),
                    document_type='transport', receipt_month=9, receipt_year=2026)

    def assert_uploaded(self, owner):
        self.cursor.fetchone.side_effect = [{'exists': 1}, self.row(owner)]
        result = self.client.post('/transport-receipts', json=self.payload(owner))
        self.assertEqual(result.status_code, 200, result.text)
        self.assertEqual(result.json()['receipt']['employee_username'], owner)
        writes = [call.args[1] for call in self.cursor.execute.call_args_list
                  if 'INSERT INTO transport_receipts' in call.args[0]]
        self.assertEqual(writes[0][:2], (owner, self.actor['username']))
        self.assertEqual(writes[0][-2:], (9, 2026))
        self.assertEqual(next(self.storage.iterdir()).read_bytes(), self.content)
        self.connection.commit.assert_called_once()

    def test_employee_can_upload_own_document(self):
        self.assert_uploaded('employee')

    def test_manager_can_upload_own_document(self):
        self.actor = {'username': 'manager', 'role_tags': ['Responsable']}
        self.assert_uploaded('manager')

    def test_admin_can_upload_on_behalf(self):
        self.actor = {'username': 'admin', 'role_tags': ['Admin', 'Responsable']}
        self.assert_uploaded('employee')

    def test_non_admin_cannot_upload_for_another_account(self):
        for tags in [['Employe'], ['Responsable'], ['Responsable', 'Employe']]:
            with self.subTest(tags=tags):
                self.actor = {'username': 'employee', 'role_tags': tags}
                payload = self.payload('team-member')
                payload['file_b64'] = 'invalid-data'
                result = self.client.post('/transport-receipts', json=payload)
                self.assertEqual(result.status_code, 403)
        self.connect.assert_not_called()
        self.assertEqual(list(self.storage.iterdir()), [])

    def test_non_admin_assignees_include_only_self(self):
        for tags in [['Employe'], ['Responsable']]:
            self.actor['role_tags'] = tags
            result = self.client.get('/transport-receipts/assignees')
            self.assertEqual(result.json(), {'users': ['employee']})
        self.connect.assert_not_called()

    def test_admin_assignees_include_all_accounts(self):
        self.actor['role_tags'] = ['Admin']
        self.cursor.fetchall.return_value = [{'username': 'admin'}, {'username': 'employee'}]
        result = self.client.get('/transport-receipts/assignees')
        self.assertEqual(result.json()['users'], ['admin', 'employee'])

    def test_unauthenticated_upload_and_assignees_are_rejected(self):
        self.actor = None
        for result in [self.client.post('/transport-receipts', json=self.payload('employee')),
                       self.client.get('/transport-receipts/assignees')]:
            self.assertEqual(result.status_code, 401)
        self.connect.assert_not_called()

    def test_employee_listing_applies_own_account_filter(self):
        self.cursor.fetchall.return_value = []
        result = self.client.get('/transport-receipts')
        self.assertEqual(result.status_code, 200)
        sql, params = self.cursor.execute.call_args.args
        self.assertIn('employee_username IN', sql)
        self.assertEqual(params, ('employee',))

    def test_manager_listing_includes_only_self_and_team(self):
        self.actor = {'username': 'manager', 'role_tags': ['Responsable']}
        self.cursor.fetchall.return_value = []
        self.assertEqual(self.client.get('/transport-receipts').status_code, 200)
        self.assertEqual(self.cursor.execute.call_args.args[1], ('manager', 'team-member'))

    def test_employee_cannot_download_another_account_document(self):
        self.cursor.fetchone.return_value = self.row('team-member')
        self.assertEqual(self.client.get('/transport-receipts/11/file').status_code, 403)

    def test_manager_team_download_and_outside_rejection(self):
        self.actor = {'username': 'manager', 'role_tags': ['Responsable']}
        (self.storage/'stored.pdf').write_bytes(self.content)
        self.cursor.fetchone.return_value = self.row('team-member')
        result = self.client.get('/transport-receipts/11/file')
        self.assertEqual(result.status_code, 200)
        self.assertEqual(result.content, self.content)
        self.cursor.fetchone.return_value = self.row('outsider')
        self.assertEqual(self.client.get('/transport-receipts/11/file').status_code, 403)

    def test_manager_does_not_gain_delete_or_archive_permission(self):
        self.actor = {'username': 'manager', 'role_tags': ['Responsable']}
        self.assertEqual(self.client.delete('/transport-receipts/11').status_code, 403)
        self.assertEqual(self.client.patch('/transport-receipts/11/archive').status_code, 403)


if __name__ == '__main__':
    unittest.main()
