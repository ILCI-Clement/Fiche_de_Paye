"""Isolated HTTP/transaction checks. No production or snapshot connections or files."""
import json
import sqlite3
import tempfile
import unittest
from contextlib import closing
from pathlib import Path
from fastapi import FastAPI, HTTPException
from fastapi.testclient import TestClient

from backend.resource_maintenance import register_resource_maintenance
from backend.account_purge import purge_account
from backend.attendance_permissions import validate_manager_save


class Cursor:
    def __init__(self, db):
        self.cursor = db.cursor()
    def __enter__(self): return self
    def __exit__(self, *_): self.cursor.close()
    def execute(self, sql, params=()):
        sql = sql.replace('%s', '?').replace(' FOR UPDATE', '').replace('INSERT IGNORE', 'INSERT OR IGNORE')
        return self.cursor.execute(sql, params)
    def fetchone(self):
        row = self.cursor.fetchone()
        return dict(row) if row else None
    def fetchall(self): return [dict(row) for row in self.cursor.fetchall()]


class Connection:
    def __init__(self, path):
        self.db = sqlite3.connect(path)
        self.db.row_factory = sqlite3.Row
    def cursor(self): return Cursor(self.db)
    def commit(self): self.db.commit()
    def rollback(self): self.db.rollback()
    def close(self): self.db.close()
    def __enter__(self): return self
    def __exit__(self, *_): self.close()


class MaintenanceTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.path = Path(self.directory.name) / 'test.sqlite'
        with closing(sqlite3.connect(self.path)) as db:
            db.executescript('''
                CREATE TABLE users (username TEXT PRIMARY KEY,email TEXT,role_tags TEXT,
                    is_active INTEGER DEFAULT 1,session_version INTEGER DEFAULT 0,
                    login_deleted_at TEXT,manager_username TEXT);
                INSERT INTO users(username,role_tags) VALUES ('super','["SuperAdmin"]'),
                    ('admin','["Admin"]'),('employee','["Employe"]');
                CREATE TABLE organization_groups(id INTEGER PRIMARY KEY,name TEXT,is_active INTEGER DEFAULT 1);
                INSERT INTO organization_groups(id,name) VALUES(1,'Department');
                CREATE TABLE user_group_memberships(username TEXT,group_id INTEGER,relation TEXT);
                CREATE TABLE leave_requests(id INTEGER PRIMARY KEY,employee_username TEXT,
                    start_date TEXT,end_date TEXT,status TEXT,is_archived INTEGER DEFAULT 0);
                INSERT INTO leave_requests VALUES(1,'employee','2026-01-01','2026-01-01','pending',0);
                INSERT INTO leave_requests VALUES(2,'employee','2026-01-01','2026-01-01','approved',0);
                CREATE TABLE leave_events(id INTEGER PRIMARY KEY,request_id INTEGER);
                INSERT INTO leave_events VALUES(1,2);
                CREATE TABLE annual_interviews(id INTEGER PRIMARY KEY,employee_username TEXT,
                    interview_year INTEGER,completed INTEGER,is_archived INTEGER DEFAULT 0);
                INSERT INTO annual_interviews VALUES(1,'employee',2026,1,0);
                CREATE TABLE permanent_deletion_events(id INTEGER PRIMARY KEY,actor_username TEXT,
                    resource_type TEXT,resource_id TEXT,metadata TEXT);
                CREATE TABLE resource_archive_events(id INTEGER PRIMARY KEY,actor_username TEXT,
                    resource_type TEXT,resource_id TEXT,is_archived INTEGER);
                CREATE TABLE personnel_contracts(id INTEGER PRIMARY KEY,username TEXT);
                CREATE TABLE transport_receipts(id INTEGER PRIMARY KEY,employee_username TEXT,uploaded_by TEXT);
                CREATE TABLE personnel_register(employee_id INTEGER PRIMARY KEY,linked_username TEXT);
                CREATE TABLE Presence(user_id TEXT,form_content TEXT);
                CREATE TABLE account_username_reservations(username TEXT PRIMARY KEY);
            ''')
        self.actor = {'username': 'admin', 'role_tags': ['Admin'], 'session_version': 0}
        self.app = FastAPI()
        register_resource_maintenance(self.app, get_db_connection=self.connect,
                                      get_current_user=lambda: self.actor, load_user=self.load_user)
        self.client = TestClient(self.app)

    def connect(self): return Connection(self.path)
    def load_user(self, cursor, username):
        cursor.execute('SELECT * FROM users WHERE username=%s', (username,))
        user = cursor.fetchone()
        if user: user['role_tags'] = json.loads(user['role_tags'])
        return user
    def sql(self, sql, params=()):
        with closing(sqlite3.connect(self.path)) as db:
            rows = db.execute(sql, params).fetchall()
            db.commit()
            return rows
    def super(self):
        self.actor = {'username': 'super', 'role_tags': ['SuperAdmin'], 'session_version': 0}
    def archive(self, kind='group', identifier=1, value=True):
        return self.client.patch(f'/maintenance/resources/{kind}/{identifier}/archive', json={'is_archived': value})
    def delete(self, kind='group', identifier=1, confirmation=None):
        return self.client.request('DELETE', f'/maintenance/resources/{kind}/{identifier}',
                                   json={'confirmation': str(identifier) if confirmation is None else confirmation})

    def test_admin_can_archive_and_restore_with_audit(self):
        self.assertEqual(self.archive().status_code, 200)
        self.assertEqual(self.client.get('/maintenance/resources/group').json()['total'], 0)
        self.assertEqual(self.client.get('/maintenance/resources/group?archive=archived').json()['total'], 1)
        self.assertEqual(self.archive(value=False).status_code, 200)
        self.assertEqual(len(self.sql('SELECT * FROM resource_archive_events')), 2)

    def test_admin_cannot_delete_even_archived(self):
        self.archive()
        self.assertEqual(self.delete().status_code, 403)
        self.assertEqual(len(self.sql('SELECT * FROM organization_groups')), 1)

    def test_employee_cannot_read_or_archive_or_delete(self):
        self.actor = {'username': 'employee', 'role_tags': ['SuperAdmin'], 'session_version': 0}
        self.assertEqual(self.client.get('/maintenance/resources/group').status_code, 403)
        self.assertEqual(self.archive().status_code, 403)
        self.assertEqual(self.delete().status_code, 403)

    def test_stale_session_cannot_archive(self):
        self.sql("UPDATE users SET session_version=1 WHERE username='admin'")
        self.assertEqual(self.archive().status_code, 403)

    def test_super_requires_archive_and_exact_confirmation(self):
        self.super()
        self.assertEqual(self.delete().status_code, 409)
        self.archive()
        self.assertEqual(self.delete(confirmation='wrong').status_code, 400)
        self.assertEqual(self.delete().status_code, 200)
        self.assertFalse(self.sql('SELECT * FROM organization_groups'))
        self.assertEqual(len(self.sql('SELECT * FROM permanent_deletion_events')), 1)

    def test_group_membership_blocks_purge(self):
        self.super(); self.archive()
        self.sql("INSERT INTO user_group_memberships VALUES('employee',1,'member')")
        self.assertEqual(self.delete().status_code, 409)

    def test_audit_failure_rolls_back_deletion(self):
        self.super(); self.archive()
        self.sql('DROP TABLE permanent_deletion_events')
        with self.assertRaises(sqlite3.OperationalError): self.delete()
        self.assertEqual(len(self.sql('SELECT * FROM organization_groups')), 1)

    def test_pending_leave_cannot_be_archived(self):
        self.assertEqual(self.archive('leave_request').status_code, 409)

    def test_processed_leave_purge_removes_only_its_history(self):
        self.super(); self.assertEqual(self.archive('leave_request', 2).status_code, 200)
        self.assertEqual(self.delete('leave_request', 2).status_code, 200)
        self.assertFalse(self.sql('SELECT * FROM leave_events WHERE request_id=2'))
        self.assertEqual(len(self.sql('SELECT * FROM leave_requests')), 1)

    def test_active_employee_interview_stays_until_account_deactivated(self):
        self.super(); self.archive('annual_interview')
        self.assertEqual(self.delete('annual_interview').status_code, 409)
        self.sql("UPDATE users SET is_active=0 WHERE username='employee'")
        self.assertEqual(self.delete('annual_interview').status_code, 200)

    def clean_employee(self):
        for table in ('annual_interviews', 'leave_requests', 'leave_events'):
            self.sql(f'DELETE FROM {table}')
        self.sql("UPDATE users SET is_active=0 WHERE username='employee'")

    def purge(self):
        db = self.connect()
        try:
            with db.cursor() as cur:
                result = purge_account(cur, {'username': 'super', 'role_tags': ['SuperAdmin']},
                                       self.load_user(cur, 'employee'), reason='Synthetic test')
            db.commit(); return result
        except Exception:
            db.rollback(); raise
        finally: db.close()

    def test_account_with_dependencies_cannot_be_physically_deleted(self):
        self.sql("UPDATE users SET is_active=0 WHERE username='employee'")
        with self.assertRaises(HTTPException) as e: self.purge()
        self.assertEqual(e.exception.status_code, 409)

    def test_empty_account_purge_reserves_old_name_and_keeps_others_files(self):
        self.clean_employee()
        self.sql("INSERT INTO transport_receipts VALUES(1,'admin','employee')")
        self.assertEqual(self.purge()['deleted_records'], 1)
        self.assertFalse(self.sql("SELECT * FROM users WHERE username='employee'"))
        self.assertEqual(self.sql('SELECT username FROM account_username_reservations'), [('employee',)])
        self.assertEqual(len(self.sql('SELECT * FROM transport_receipts')), 1)

    def test_legacy_email_attendance_blocks_account_purge(self):
        self.clean_employee()
        self.sql("UPDATE users SET email='employee@example.test' WHERE username='employee'")
        self.sql('INSERT INTO Presence VALUES(?,?)', ('admin', json.dumps({'employes_data': [
            {'id': 'legacy', 'email_employe': 'EMPLOYEE@example.test'}]})))
        with self.assertRaises(HTTPException) as e: self.purge()
        self.assertEqual(e.exception.status_code, 409)

    def test_admin_cannot_remove_a_sheet_via_workspace_save(self):
        with self.connect() as db, db.cursor() as cursor:
            with self.assertRaises(HTTPException) as e:
                validate_manager_save(cursor, self.actor, {'employes_data': [{'id': '1', 'is_archived': True}]},
                                      {'employes_data': []}, self.load_user, lambda *_: True, confirmed_deletions=['1'])
        self.assertEqual(e.exception.status_code, 403)

    def test_super_sheet_delete_requires_explicit_ids_and_archival(self):
        self.super()
        with self.connect() as db, db.cursor() as cursor:
            for archived, ids, status in [(False, ['1'], 409), (True, [], 400)]:
                with self.assertRaises(HTTPException) as e:
                    validate_manager_save(cursor, self.actor, {'employes_data': [{'id': '1', 'is_archived': archived}]},
                                          {'employes_data': []}, self.load_user, lambda *_: True, confirmed_deletions=ids)
                self.assertEqual(e.exception.status_code, status)
            validate_manager_save(cursor, self.actor, {'employes_data': [{'id': '1', 'is_archived': True}]},
                                  {'employes_data': []}, self.load_user, lambda *_: True, confirmed_deletions=['1'])

    def test_admin_archive_retains_sheet_and_audits(self):
        db = self.connect()
        with db.cursor() as cursor:
            validate_manager_save(cursor, self.actor, {'employes_data': [{'id': '1'}]},
                                  {'employes_data': [{'id': '1', 'is_archived': True}]}, self.load_user, lambda *_: True)
        db.commit(); db.close()
        self.assertEqual(len(self.sql('SELECT * FROM resource_archive_events')), 1)


if __name__ == '__main__': unittest.main()
