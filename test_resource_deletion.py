"""Archived-resource deletion against temporary SQLite/files, never live personnel."""
from contextlib import closing
import json
from pathlib import Path
import sqlite3
import tempfile
import unittest
from unittest.mock import patch

from fastapi import FastAPI, HTTPException
from fastapi.testclient import TestClient

from backend.contract_history import register_contract_history
from backend.documents import register_document_routes
from backend.personnel_register import register_personnel_routes
from backend.role_policy import canonical_role_tags, is_admin
import test_personnel_register_sensitive as sensitive_fixture


class ResourceCursor(sensitive_fixture.Cursor):
    def execute(self, sql, params=()):
        if self.connection.fixture.fail_deletion_audit and 'INSERT INTO permanent_deletion_events' in sql:
            raise RuntimeError('Injected isolated deletion audit failure')
        return super().execute(sql, params)


class ResourceConnection(sensitive_fixture.Connection):
    def cursor(self): return ResourceCursor(self)


class ResourceDeletionTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.path = Path(self.temp.name) / 'resources.sqlite3'
        self.storage = Path(self.temp.name) / 'files'
        self.storage.mkdir()
        self.file = self.storage / 'fixture.pdf'
        self.file.write_bytes(b'%PDF-1.4 isolated document fixture')
        self.fail_audit = False
        self.fail_deletion_audit = False
        with closing(sqlite3.connect(self.path)) as db, db:
            db.executescript('''CREATE TABLE users (
                username TEXT PRIMARY KEY,email TEXT,role TEXT,role_tags TEXT,is_admin INTEGER DEFAULT 0,
                is_active INTEGER DEFAULT 1,session_version INTEGER DEFAULT 0,login_deleted_at TEXT,
                contract_start_date TEXT,contract_end_date TEXT,is_cdi INTEGER DEFAULT 0);
                CREATE TABLE personnel_register (
                employee_id INTEGER PRIMARY KEY AUTOINCREMENT,full_name TEXT NOT NULL,
                nationality TEXT DEFAULT '',sex TEXT DEFAULT '',address TEXT DEFAULT '',birth_date TEXT,
                job_title TEXT DEFAULT '',qualification TEXT DEFAULT '',personnel_type TEXT DEFAULT '',
                linked_username TEXT UNIQUE,social_security_number_encrypted TEXT,
                is_archived INTEGER DEFAULT 0,version INTEGER DEFAULT 1,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP,updated_at TEXT DEFAULT CURRENT_TIMESTAMP);
                CREATE TABLE personnel_register_periods (
                period_id INTEGER PRIMARY KEY AUTOINCREMENT,employee_id INTEGER,
                entry_date TEXT,exit_date TEXT,contract_type TEXT DEFAULT '',other_contract TEXT DEFAULT '',
                created_at TEXT DEFAULT CURRENT_TIMESTAMP,updated_at TEXT DEFAULT CURRENT_TIMESTAMP);
                CREATE TABLE personnel_register_events (
                event_id INTEGER PRIMARY KEY AUTOINCREMENT,employee_id INTEGER,actor_username TEXT,
                action TEXT,before_snapshot TEXT,after_snapshot TEXT,created_at TEXT DEFAULT CURRENT_TIMESTAMP);
                CREATE TABLE permanent_deletion_events (
                id INTEGER PRIMARY KEY AUTOINCREMENT,actor_username TEXT,resource_type TEXT,
                resource_id TEXT,metadata TEXT,created_at TEXT DEFAULT CURRENT_TIMESTAMP);
                CREATE TABLE personnel_contracts (
                id INTEGER PRIMARY KEY AUTOINCREMENT,username TEXT,employee_type TEXT,
                start_date TEXT,end_date TEXT,is_cdi INTEGER DEFAULT 0,is_archived INTEGER DEFAULT 0,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP);
                CREATE TABLE Presence (user_id TEXT,form_content TEXT);
                CREATE TABLE transport_receipts (
                id INTEGER PRIMARY KEY AUTOINCREMENT,employee_username TEXT,uploaded_by TEXT,
                original_filename TEXT,stored_filename TEXT,mime_type TEXT,size_bytes INTEGER,
                document_type TEXT,other_type TEXT,receipt_month INTEGER,receipt_year INTEGER,
                document_number TEXT,archived INTEGER DEFAULT 0,created_at TEXT DEFAULT CURRENT_TIMESTAMP,
                archived_at TEXT,archived_by TEXT);''')
            for name, role in [('fixture-super', 'SuperAdmin'), ('fixture-admin', 'Admin'),
                               ('fixture-manager', 'Responsable'), ('fixture-employee', 'Employe')]:
                db.execute('INSERT INTO users(username,email,role,role_tags,is_admin) VALUES(?,?,?,?,?)',
                           (name, name + '@example.test', role, json.dumps([role]), role == 'Admin'))
            db.execute('''INSERT INTO transport_receipts
                (employee_username,uploaded_by,original_filename,stored_filename,mime_type,size_bytes,
                document_type,receipt_month,receipt_year,document_number)
                VALUES(?,?,?,?,?,?,?,?,?,?)''', ('fixture-employee', 'fixture-admin', 'fixture.pdf',
                    'fixture.pdf', 'application/pdf', self.file.stat().st_size, 'transport', 10, 2026, 'fixture-1'))
            db.execute('''INSERT INTO personnel_contracts(username,employee_type,start_date,end_date,is_cdi)
                VALUES(?,?,?,?,?)''', ('fixture-employee', 'salarie', '2020-01-01', None, 1))
        self.actor = self.identity('fixture-admin', 'Admin')

        def authenticated():
            if self.actor is None:
                raise HTTPException(401, 'Authentication required')
            return dict(self.actor)

        def load_user(cursor, username):
            cursor.execute('SELECT * FROM users WHERE username=%s', (username,))
            row = cursor.fetchone()
            if row:
                row['role_tags'] = canonical_role_tags(json.loads(row['role_tags']))
            return row

        def require_roles(*roles):
            def authorize():
                actor = authenticated()
                if not set(roles).intersection(canonical_role_tags(actor['role_tags'])):
                    raise HTTPException(403, 'Forbidden')
                return actor
            return authorize

        def access(cursor, actor, username, edit=False):
            if not is_admin(actor) and actor['username'] != username:
                raise HTTPException(403, 'Forbidden')

        app = FastAPI()
        connect = lambda: ResourceConnection(self)
        register_personnel_routes(app, get_db_connection=connect, get_current_user=authenticated, load_user=load_user)
        register_document_routes(app, get_db_connection=connect, get_current_user=authenticated,
            require_roles=require_roles, load_user=load_user, user_can_manage=lambda *_: False,
            all_users=lambda *_: [], storage_dir=self.storage, max_file_bytes=1024 * 1024)
        register_contract_history(app, connect, require_roles, authenticated, access, load_user=load_user)
        self.client = TestClient(app, raise_server_exceptions=False)
        self.addCleanup(self.client.close)

    @staticmethod
    def identity(name, role):
        return {'username': name, 'role_tags': [role], 'session_version': 0}

    def super_actor(self): self.actor = self.identity('fixture-super', 'SuperAdmin')

    def sql(self, sql, params=()):
        with closing(sqlite3.connect(self.path)) as db, db:
            db.row_factory = sqlite3.Row
            cursor = db.execute(sql, params)
            return [dict(row) for row in cursor.fetchall()]

    def ledger(self): return self.sql('SELECT * FROM permanent_deletion_events ORDER BY id')

    def personnel(self, archived=False):
        response = self.client.post('/personnel-register', json={
            'full_name': 'Synthetic employee identity', 'address': 'Synthetic private address',
            'initial_period': {'entry_date': '2020-01-01'},
        })
        self.assertEqual(response.status_code, 201, response.text)
        employee = response.json()['employee']
        if archived:
            response = self.client.patch(f"/personnel-register/{employee['employee_id']}/archive", json={
                'is_archived': True, 'expected_version': employee['version'],
            })
            self.assertEqual(response.status_code, 200, response.text)
            employee = response.json()['employee']
        return employee

    def delete_personnel(self, employee, confirmation=None, version=None):
        return self.client.request('DELETE', f"/personnel-register/{employee['employee_id']}", json={
            'confirmation': confirmation if confirmation is not None else str(employee['employee_id']),
            'expected_version': version if version is not None else employee['version'],
        })

    def archive_document(self, archived=True):
        return self.client.patch('/transport-receipts/1/archive', json={'is_archived': archived})

    def delete_document(self, confirmation='1'):
        return self.client.request('DELETE', '/transport-receipts/1', json={'confirmation': confirmation})

    def archive_contract(self, archived=True):
        return self.client.patch('/users/fixture-employee/contracts/1/archive', json={'is_archived': archived})

    def delete_contract(self, confirmation='1'):
        return self.client.request('DELETE', '/users/fixture-employee/contracts/1', json={'confirmation': confirmation})

    def test_admin_retains_create_edit_archive_but_cannot_delete_registry(self):
        employee = self.personnel(archived=True)
        self.assertEqual(self.delete_personnel(employee).status_code, 403)
        self.assertEqual(len(self.sql('SELECT * FROM personnel_register')), 1)
        self.assertEqual(self.ledger(), [])

    def test_super_deletion_requires_archive_confirmation_and_current_version(self):
        employee = self.personnel()
        self.super_actor()
        self.assertEqual(self.delete_personnel(employee).status_code, 409)
        response = self.client.patch(f"/personnel-register/{employee['employee_id']}/archive", json={
            'is_archived': True, 'expected_version': employee['version'],
        })
        self.assertEqual(response.status_code, 200, response.text)
        archived = response.json()['employee']
        self.assertEqual(self.delete_personnel(archived, confirmation='wrong').status_code, 400)
        self.assertEqual(self.delete_personnel(archived, version=employee['version']).status_code, 409)
        self.assertEqual(self.delete_personnel(archived, confirmation=archived['full_name']).status_code, 200)

    def test_registry_deletion_removes_local_pii_history_but_retains_global_metadata(self):
        employee = self.personnel(archived=True)
        self.sql('UPDATE personnel_register SET social_security_number_encrypted=? WHERE employee_id=?',
                 ('synthetic-envelope', employee['employee_id']))
        self.super_actor()
        self.assertEqual(self.delete_personnel(employee).status_code, 200)
        for table in ('personnel_register', 'personnel_register_periods', 'personnel_register_events'):
            self.assertEqual(self.sql('SELECT * FROM ' + table), [])
        event = self.ledger()[0]
        self.assertEqual(event['resource_type'], 'personnel')
        self.assertEqual(json.loads(event['metadata'])['period_count'], 1)
        self.assertNotIn('Synthetic', json.dumps(event))
        self.assertNotIn('envelope', json.dumps(event))
        self.assertEqual(len(self.sql('SELECT * FROM users')), 4)
        self.assertEqual(len(self.sql('SELECT * FROM personnel_contracts')), 1)
        self.assertTrue(self.file.exists())

    def test_registry_deletion_ledger_failure_rolls_back_all_pii_rows(self):
        employee = self.personnel(archived=True)
        before = [self.sql('SELECT * FROM ' + table) for table in
                  ('personnel_register', 'personnel_register_periods', 'personnel_register_events')]
        self.super_actor(); self.fail_deletion_audit = True
        with self.assertLogs('backend.personnel_register', level='ERROR'):
            response = self.delete_personnel(employee)
        self.assertEqual(response.status_code, 500)
        self.assertEqual(before, [self.sql('SELECT * FROM ' + table) for table in
                         ('personnel_register', 'personnel_register_periods', 'personnel_register_events')])

    def test_employee_and_admin_can_archive_own_documents_but_never_delete(self):
        for name, role in [('fixture-employee', 'Employe'), ('fixture-admin', 'Admin')]:
            self.actor = self.identity(name, role)
            self.assertEqual(self.archive_document(True).status_code, 200)
            self.assertEqual(self.archive_document(False).status_code, 200)
            self.assertEqual(self.delete_document().status_code, 403)
        self.assertTrue(self.file.exists())

    def test_manager_remains_read_only_even_for_own_upload(self):
        self.sql('UPDATE transport_receipts SET employee_username=?,uploaded_by=? WHERE id=1',
                 ('fixture-manager', 'fixture-manager'))
        self.actor = self.identity('fixture-manager', 'Responsable')
        self.assertEqual(self.archive_document().status_code, 403)
        self.assertEqual(self.delete_document().status_code, 403)
        self.assertTrue(self.file.exists())

    def test_super_document_delete_requires_archive_confirmation_then_removes_file(self):
        self.super_actor()
        self.assertEqual(self.delete_document().status_code, 409)
        self.assertEqual(self.archive_document().status_code, 200)
        self.assertEqual(self.delete_document('wrong').status_code, 400)
        response = self.delete_document()
        self.assertEqual(response.status_code, 200, response.text)
        self.assertFalse(response.json()['cleanup_pending'])
        self.assertFalse(self.file.exists())
        self.assertEqual(self.sql('SELECT * FROM transport_receipts'), [])
        self.assertEqual(self.ledger()[0]['resource_type'], 'document')
        self.assertNotIn('fixture.pdf', self.ledger()[0]['metadata'])

    def test_document_audit_failure_keeps_database_row_and_file(self):
        self.super_actor(); self.archive_document(); self.fail_deletion_audit = True
        self.assertEqual(self.delete_document().status_code, 500)
        self.assertEqual(len(self.sql('SELECT * FROM transport_receipts')), 1)
        self.assertTrue(self.file.exists())
        self.assertEqual(self.ledger(), [])

    def test_failed_file_cleanup_is_reported_without_false_physical_deletion(self):
        self.super_actor(); self.archive_document()
        with patch.object(Path, 'unlink', side_effect=PermissionError('Synthetic failure')):
            response = self.delete_document()
        self.assertEqual(response.status_code, 200, response.text)
        self.assertTrue(response.json()['cleanup_pending'])
        self.assertEqual(response.json()['status'], 'cleanup_pending')
        self.assertTrue(self.file.exists())
        self.assertEqual(self.sql('SELECT * FROM transport_receipts'), [])
        self.assertNotIn('deleted_files', json.loads(self.ledger()[0]['metadata']))

    def test_contract_archive_retains_history_and_admin_cannot_delete(self):
        self.assertEqual(self.archive_contract().status_code, 200)
        result = self.client.get('/users/fixture-employee/contracts')
        self.assertEqual(result.status_code, 200, result.text)
        self.assertTrue(result.json()['contracts'][0]['is_archived'])
        self.assertEqual(self.delete_contract().status_code, 403)
        self.assertEqual(self.archive_contract(False).status_code, 200)

    def test_super_contract_delete_requires_archived_and_confirmation(self):
        self.super_actor()
        self.assertEqual(self.delete_contract().status_code, 409)
        self.archive_contract()
        self.assertEqual(self.delete_contract('wrong').status_code, 400)
        response = self.delete_contract()
        self.assertEqual(response.status_code, 200, response.text)
        self.assertEqual(self.sql('SELECT * FROM personnel_contracts'), [])
        self.assertEqual(self.ledger()[0]['resource_type'], 'contract')

    def test_contract_referenced_in_another_workspace_cannot_be_deleted(self):
        self.super_actor(); self.archive_contract()
        for reference in ({'contract_id': 1, 'account_username': 'fixture-employee'},
                          {'contract_snapshot': {'id': 1, 'username': 'fixture-employee'}},
                          {'contract_id': 1}):
            self.sql('DELETE FROM Presence')
            self.sql('INSERT INTO Presence VALUES(?,?)',
                     ('another-workspace-owner', json.dumps({'employes_data': [reference]})))
            self.assertEqual(self.delete_contract().status_code, 409)
            self.assertEqual(len(self.sql('SELECT * FROM personnel_contracts')), 1)
        self.assertEqual(self.ledger(), [])

    def test_contract_reference_matches_both_id_and_employee(self):
        self.super_actor(); self.archive_contract()
        self.sql('INSERT INTO Presence VALUES(?,?)', ('another-workspace', json.dumps({
            'employes_data': [{'contract_id': 1, 'account_username': 'another-employee'}],
        })))
        self.assertEqual(self.delete_contract().status_code, 200)
        self.assertEqual(len(self.sql('SELECT * FROM Presence')), 1)

    def test_contract_delete_refuses_unreadable_attendance_without_data_loss(self):
        self.super_actor(); self.archive_contract()
        self.sql('INSERT INTO Presence VALUES(?,?)', ('another-workspace', '{bad-json'))
        self.assertEqual(self.delete_contract().status_code, 409)
        self.assertEqual(len(self.sql('SELECT * FROM personnel_contracts')), 1)

    def test_archived_contract_does_not_block_creation_but_restore_checks_overlap(self):
        self.archive_contract()
        response = self.client.post('/users/fixture-employee/contracts', json={
            'employee_type': 'salarie', 'start_date': '2020-01-01', 'is_cdi': True,
        })
        self.assertEqual(response.status_code, 200, response.text)
        self.assertEqual(len(self.sql('SELECT * FROM personnel_contracts')), 2)
        self.assertEqual(self.archive_contract(False).status_code, 409)
        self.assertTrue(self.sql('SELECT * FROM personnel_contracts WHERE id=1')[0]['is_archived'])

    def test_contract_ledger_failure_rolls_back_deletion(self):
        self.super_actor(); self.archive_contract(); self.fail_deletion_audit = True
        self.assertEqual(self.delete_contract().status_code, 500)
        self.assertEqual(len(self.sql('SELECT * FROM personnel_contracts')), 1)
        self.assertEqual(self.ledger(), [])

    def test_revoked_super_role_and_session_are_rechecked_inside_writes(self):
        employee = self.personnel(archived=True)
        self.super_actor(); self.archive_document(); self.archive_contract()
        self.sql('UPDATE users SET role_tags=? WHERE username=?', ('["Admin"]', 'fixture-super'))
        for response in (self.delete_personnel(employee), self.delete_document(), self.delete_contract()):
            self.assertEqual(response.status_code, 403, response.text)
        self.sql('UPDATE users SET session_version=1 WHERE username=?', ('fixture-super',))
        for response in (self.delete_personnel(employee), self.delete_document(), self.delete_contract()):
            self.assertEqual(response.status_code, 401, response.text)
        self.assertEqual(self.ledger(), [])


if __name__ == '__main__':
    unittest.main()
