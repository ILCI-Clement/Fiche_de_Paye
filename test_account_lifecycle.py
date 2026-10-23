"""Isolated HTTP/auth/transaction regression; no live accounts, database or files."""
import hashlib
import json
from pathlib import Path
import secrets
import sqlite3
import tempfile
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor
from contextlib import closing
import unittest
from unittest.mock import patch

import test_backend_helpers  # isolated environment defaults, no database connection
from backend import main
from backend.account_lifecycle import register_account_lifecycle, ensure_account_lifecycle_schema
from fastapi import Depends, FastAPI
from fastapi.testclient import TestClient


def digest(value):
    return hashlib.sha256(value.encode()).hexdigest()


class Cursor:
    def __init__(self, connection):
        self.connection = connection
        self.cursor = connection.db.cursor()

    def __enter__(self): return self
    def __exit__(self, *_): self.cursor.close()

    @property
    def description(self): return self.cursor.description

    def execute(self, sql, params=()):
        if sql == "SHOW TABLES LIKE 'personnel_register'":
            return self.cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='personnel_register'")
        if 'FOR UPDATE' in sql:
            # SQLite surrogate for transaction serialization, not a MySQL lock test.
            if not self.connection.db.in_transaction:
                self.connection.db.execute('BEGIN IMMEDIATE')
            sql = sql.replace(' FOR UPDATE', '')
        if self.connection.fixture.fail_event and 'INSERT INTO account_status_events' in sql:
            raise RuntimeError('Injected audit write failure')
        return self.cursor.execute(sql.replace('%s', '?'), params)

    def fetchone(self):
        row = self.cursor.fetchone()
        if row is None:
            return None
        value = dict(row)
        for key in ('login_now', 'login_locked_until', 'login_last_failed_at'):
            if value.get(key):
                value[key] = datetime.fromisoformat(value[key])
        return value

    def fetchall(self): return [dict(row) for row in self.cursor.fetchall()]


class Connection:
    def __init__(self, fixture):
        self.fixture = fixture
        self.db = sqlite3.connect(fixture.path, timeout=10, isolation_level=None)
        self.db.row_factory = sqlite3.Row
        self.db.create_function('UTC_TIMESTAMP', 1, lambda _: datetime.utcnow().isoformat(' '))

    def cursor(self): return Cursor(self)
    def commit(self): self.db.commit()
    def rollback(self): self.db.rollback()
    def close(self): self.db.close()


class AccountLifecycleTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.path = Path(self.temp.name) / 'fixture.sqlite3'
        self.fail_event = False
        self.password = secrets.token_urlsafe(18)
        self.invoice = Path(self.temp.name) / 'invoice.pdf'
        self.invoice.write_bytes(b'%PDF-1.4 isolated retention fixture')
        with closing(sqlite3.connect(self.path)) as db, db:
            db.executescript('''CREATE TABLE users (
                username TEXT PRIMARY KEY, email TEXT UNIQUE, role_tags TEXT, password_hash TEXT,
                is_active INTEGER DEFAULT 1, session_version INTEGER DEFAULT 0, login_deleted_at TEXT,
                login_failed_attempts INTEGER NOT NULL DEFAULT 0,login_locked_until TEXT,login_last_failed_at TEXT,
                remember_token_hash TEXT,remember_token_expires TEXT,reset_token TEXT,reset_token_expires TEXT);
                CREATE TABLE account_status_events (id INTEGER PRIMARY KEY AUTOINCREMENT,
                target_username TEXT,actor_username TEXT,was_active INTEGER,is_active INTEGER,
                action TEXT DEFAULT 'activation',created_at TEXT DEFAULT CURRENT_TIMESTAMP);
                CREATE TABLE retained_business (kind TEXT,owner TEXT,uploaded_by TEXT,archived_by TEXT,payload TEXT);''')
            for username, tags in [('admin', ['SuperAdmin', 'Admin']), ('other-admin', ['Admin']), ('manager', ['Responsable']), ('employee', ['Employe'])]:
                db.execute('INSERT INTO users(username,email,role_tags,password_hash) VALUES(?,?,?,?)',
                           (username, username + '@example.test', json.dumps(tags), digest(self.password)))
            for kind, owner in [('attendance', 'employee'), ('contracts', 'employee'), ('documents', 'employee'),
                                ('documents', 'other-person'), ('leave', 'employee'), ('interviews', 'employee')]:
                db.execute('INSERT INTO retained_business VALUES(?,?,?,?,?)',
                           (kind, owner, 'employee', 'employee', 'unchanged historical payload'))
        self.addCleanup(patch.stopall)
        patch.object(main, 'get_db_connection', self.connect).start()
        patch.object(main, 'load_user', self.load_user).start()
        patch.object(main, 'verify_password', lambda value, hashed: digest(value) == hashed).start()
        app = FastAPI()
        register_account_lifecycle(app, get_db_connection=self.connect, get_current_user=main.get_current_user,
                                   load_user=self.load_user, public_user=main.public_user, hash_password=digest)
        app.add_api_route('/login', main.login, methods=['POST'])
        app.add_api_route('/restore-session', main.restore_session, methods=['POST'])
        app.add_api_route('/delete-user/{username_to_delete}', main.delete_user, methods=['DELETE'])

        @app.get('/private')
        def private(actor=Depends(main.get_current_user)):
            return {'username': actor['username']}

        self.client = TestClient(app, raise_server_exceptions=False)
        self.addCleanup(self.client.close)

    def connect(self): return Connection(self)

    def load_user(self, cursor, username):
        cursor.execute('SELECT * FROM users WHERE username=%s', (username,))
        row = cursor.fetchone()
        if row:
            tags = json.loads(row['role_tags'])
            row.update(role_tags=tags, role=tags[0], is_admin='Admin' in tags)
        return row

    def stored(self, username='employee'):
        db = self.connect()
        try:
            with db.cursor() as cursor:
                return self.load_user(cursor, username)
        finally:
            db.close()

    def headers(self, username='admin', version=0):
        return {'Authorization': 'Bearer ' + main.create_session_token(username, version)}

    def activation(self, active, username='employee', actor='admin'):
        return self.client.patch(f'/users/{username}/activation', json={'is_active': active}, headers=self.headers(actor))

    def delete(self, username='employee', actor='admin', confirmation=None):
        return self.client.request('DELETE', f'/users/{username}/login', headers=self.headers(actor),
                                   json={'confirmation': username if confirmation is None else confirmation})

    def login(self, username='employee', remember=False):
        return self.client.post('/login', json={'username': username, 'password': self.password, 'remember_me': remember})

    def events(self):
        return self.client.get('/users/employee/activation-history', headers=self.headers()).json()['events']

    def test_login_deletion_accepts_exact_email_confirmation(self):
        self.assertEqual(self.activation(False).status_code, 200)
        response = self.client.request('DELETE', '/users/employee/login', headers=self.headers(),
                                       json={'confirmation': 'employee@example.test'})
        self.assertEqual(response.status_code, 200)
        self.assertEqual(self.events()[0]['action'], 'login_deleted')

    def test_login_deletion_does_not_treat_person_label_as_confirmation(self):
        self.assertEqual(self.activation(False).status_code, 200)
        for label in ['Jean Martin', 'EMPLOYEE@EXAMPLE.TEST', 'employee@example.test ']:
            response = self.client.request('DELETE', '/users/employee/login', headers=self.headers(),
                                           json={'confirmation': label})
            self.assertEqual(response.status_code, 400)
            self.assertIsNone(self.stored()['login_deleted_at'])

    def retained(self):
        with closing(sqlite3.connect(self.path)) as db:
            return db.execute('SELECT * FROM retained_business ORDER BY kind,owner').fetchall(), self.invoice.read_bytes()

    def test_email_login_uses_canonical_identity_and_remember_token(self):
        response = self.login('  EMPLOYEE@EXAMPLE.TEST  ', remember=True)
        self.assertEqual(response.status_code, 200)
        data = response.json()
        self.assertEqual(main.read_session_token(data['auth_token']), 'employee')
        self.assertEqual(self.client.get('/private', headers={'Authorization': 'Bearer ' + data['auth_token']}).json()['username'], 'employee')
        restored = self.client.post('/restore-session', json={'remember_token': data['remember_token']})
        self.assertEqual(restored.status_code, 200)
        self.assertEqual(main.read_session_token(restored.json()['auth_token']), 'employee')

    def test_username_login_is_case_insensitive_and_trimmed(self):
        self.assertEqual(self.login(' EMPLOYEE ').status_code, 200)

    def test_email_login_rejects_wrong_password_and_unknown_or_empty_identifier(self):
        response = self.client.post('/login', json={'username': 'employee@example.test', 'password': self.password + 'x'})
        self.assertEqual(response.status_code, 401)
        for identifier in ['', '  ', 'missing@example.test', "' OR 1=1 --"]:
            self.assertEqual(self.login(identifier).status_code, 401)

    def test_email_login_rejects_username_email_collision(self):
        with closing(sqlite3.connect(self.path)) as db, db:
            db.execute('UPDATE users SET username=? WHERE username=?', ('employee@example.test', 'manager'))
        self.assertEqual(self.login('employee@example.test').status_code, 401)

    def test_email_login_rejects_case_variant_duplicate_addresses(self):
        with closing(sqlite3.connect(self.path)) as db, db:
            db.execute('UPDATE users SET email=? WHERE username=?', ('EMPLOYEE@example.test', 'manager'))
        self.assertEqual(self.login('employee@example.test').status_code, 401)

    def test_email_login_cannot_bypass_inactive_or_deleted_status(self):
        self.assertEqual(self.activation(False).status_code, 200)
        self.assertEqual(self.login('employee@example.test').status_code, 401)
        self.assertEqual(self.delete().status_code, 200)
        self.assertEqual(self.login('employee@example.test').status_code, 401)

    def test_deletion_requires_deactivation_first(self):
        self.assertEqual(self.delete().status_code, 409)
        self.assertIsNone(self.stored()['login_deleted_at'])
        self.assertEqual(self.events(), [])

    def test_deactivate_then_delete_preserves_all_history_and_other_person_files(self):
        before = self.retained()
        self.assertEqual(self.activation(False).status_code, 200)
        self.assertEqual(self.delete().status_code, 200)
        self.assertEqual(self.retained(), before)
        row = self.stored()
        self.assertEqual(row['email'], 'employee@example.test')
        self.assertEqual(row['role_tags'], ['Employe'])
        self.assertIsNotNone(row['login_deleted_at'])
        self.assertNotEqual(row['password_hash'], digest(self.password))
        self.assertEqual(self.events()[0]['action'], 'login_deleted')

    def test_inactive_login_and_existing_session_are_rejected(self):
        token = self.login().json()['auth_token']
        self.assertEqual(self.client.get('/private', headers={'Authorization': 'Bearer ' + token}).status_code, 200)
        self.activation(False)
        self.assertEqual(self.login().status_code, 401)
        self.assertEqual(self.client.get('/private', headers={'Authorization': 'Bearer ' + token}).status_code, 401)

    def test_pre_migration_token_is_accepted_only_until_first_state_change(self):
        payload = main.encode_segment(json.dumps({'username': 'employee', 'expires_at': int(main.time.time()) + 300}).encode())
        signature = main.hmac.new(main.API_TOKEN.encode(), payload.encode(), main.hashlib.sha256).digest()
        headers = {'Authorization': 'Bearer ' + payload + '.' + main.encode_segment(signature)}
        self.assertEqual(self.client.get('/private', headers=headers).status_code, 200)
        self.activation(False); self.activation(True)
        self.assertEqual(self.client.get('/private', headers=headers).status_code, 401)

    def test_wrong_version_even_with_valid_signature_is_rejected(self):
        self.assertEqual(self.client.get('/private', headers=self.headers('employee', 999)).status_code, 401)

    def test_reactivation_does_not_revive_old_session(self):
        old = self.headers('employee')
        self.activation(False)
        self.assertEqual(self.activation(True).status_code, 200)
        self.assertEqual(self.client.get('/private', headers=old).status_code, 401)
        fresh = self.login().json()['auth_token']
        self.assertEqual(self.client.get('/private', headers={'Authorization': 'Bearer ' + fresh}).status_code, 200)

    def test_deleted_login_cannot_be_reactivated(self):
        self.activation(False); self.delete()
        self.assertEqual(self.activation(True).status_code, 409)
        self.assertEqual(self.activation(False).status_code, 409)
        self.assertEqual(self.login().status_code, 401)

    def test_remembered_login_and_reset_tokens_are_revoked(self):
        remembered = self.login(remember=True).json()['remember_token']
        with closing(sqlite3.connect(self.path)) as db, db:
            db.execute("UPDATE users SET reset_token='temporary-fixture' WHERE username='employee'")
        self.activation(False)
        self.assertIsNone(self.stored()['remember_token_hash'])
        self.assertIsNone(self.stored()['reset_token'])
        self.assertEqual(self.client.post('/restore-session', json={'remember_token': remembered}).status_code, 401)
        self.activation(True)
        self.assertEqual(self.client.post('/restore-session', json={'remember_token': remembered}).status_code, 401)

    def test_only_admin_may_modify_account_access(self):
        for actor in ['employee', 'manager']:
            self.assertEqual(self.activation(False, actor=actor).status_code, 403)
            self.assertEqual(self.delete(actor=actor).status_code, 403)
            self.assertEqual(self.client.get('/users/employee/activation-history', headers=self.headers(actor)).status_code, 403)

    def test_ordinary_admin_may_archive_but_cannot_delete_login(self):
        self.assertEqual(self.activation(False, actor='other-admin').status_code, 200)
        self.assertEqual(self.delete(actor='other-admin').status_code, 403)
        self.assertIsNone(self.stored()['login_deleted_at'])

    def test_unauthenticated_access_rejected(self):
        self.assertEqual(self.client.patch('/users/employee/activation', json={'is_active': False}).status_code, 401)
        self.assertEqual(self.client.request('DELETE', '/users/employee/login', json={'confirmation': 'employee'}).status_code, 401)

    def test_self_deactivation_and_self_deletion_rejected(self):
        self.assertEqual(self.activation(False, username='admin').status_code, 400)
        self.assertEqual(self.delete(username='admin').status_code, 400)

    def test_exact_username_confirmation_required(self):
        self.activation(False)
        for confirmation in ['', 'Employee', 'other-person']:
            self.assertEqual(self.delete(confirmation=confirmation).status_code, 400)
        self.assertIsNone(self.stored()['login_deleted_at'])

    def test_strict_boolean_and_missing_confirmation_rejected(self):
        self.assertEqual(self.client.patch('/users/employee/activation', headers=self.headers(), json={'is_active': 'false'}).status_code, 422)
        self.assertEqual(self.client.request('DELETE', '/users/employee/login', headers=self.headers(), json={}).status_code, 422)

    def test_repeated_requests_are_idempotent(self):
        self.activation(False); self.activation(False)
        self.assertEqual(self.stored()['session_version'], 1)
        self.delete(); self.delete()
        self.assertEqual(self.stored()['session_version'], 2)
        self.assertEqual(len(self.events()), 2)

    def test_unknown_account_not_found(self):
        self.assertEqual(self.activation(False, username='missing').status_code, 404)
        self.assertEqual(self.delete(username='missing').status_code, 404)

    def test_destructive_endpoint_requires_super_role_and_explicit_payload(self):
        before = self.retained()
        self.assertEqual(self.client.delete('/delete-user/employee', headers=self.headers('other-admin')).status_code, 403)
        self.assertEqual(self.client.delete('/delete-user/employee', headers=self.headers()).status_code, 422)
        self.assertEqual(self.retained(), before)

    def test_deleted_username_and_email_stay_reserved(self):
        self.activation(False); self.delete()
        with closing(sqlite3.connect(self.path)) as db, db:
            for username, email in [('employee', 'new@example.test'), ('new', 'employee@example.test')]:
                with self.assertRaises(sqlite3.IntegrityError):
                    db.execute('INSERT INTO users(username,email) VALUES(?,?)', (username, email))

    def test_audit_failure_rolls_back_deactivation(self):
        before = self.stored()
        self.fail_event = True
        self.assertEqual(self.activation(False).status_code, 500)
        self.assertEqual(self.stored(), before)

    def test_audit_failure_rolls_back_credential_erasure(self):
        self.activation(False)
        before = self.stored()
        self.fail_event = True
        self.assertEqual(self.delete().status_code, 500)
        self.assertEqual(self.stored(), before)

    def test_deletion_racing_reactivation_cannot_both_succeed(self):
        self.activation(False)
        with ThreadPoolExecutor(max_workers=2) as pool:
            results = list(pool.map(lambda operation: operation().status_code,
                                    [self.delete, lambda: self.activation(True)]))
        self.assertEqual(sorted(results), [200, 409])
        row = self.stored()
        self.assertFalse(row['is_active'] and row['login_deleted_at'])

    def test_stale_or_demoted_admin_session_is_rejected(self):
        self.activation(False, username='other-admin')
        self.activation(True, username='other-admin')
        self.assertEqual(self.activation(False, actor='other-admin').status_code, 401)
        with closing(sqlite3.connect(self.path)) as db, db:
            db.execute("UPDATE users SET role_tags=? WHERE username='other-admin'", (json.dumps(['Employe']),))
        response = self.client.patch('/users/employee/activation', json={'is_active': False}, headers=self.headers('other-admin', 2))
        self.assertEqual(response.status_code, 403)


class MigrationCursor:
    def __init__(self, fail=False, existing=False):
        self.fail = fail
        self.existing = existing
        self.sql = []
        self.last = ''

    def execute(self, sql, params=()):
        self.sql.append((sql, params)); self.last = sql
        if self.fail and sql.startswith('ALTER TABLE users'):
            raise RuntimeError('Injected DDL failure')

    def fetchone(self): return {'value': 'STRICT_TRANS_TABLES,NO_ZERO_DATE,NO_ZERO_IN_DATE'}
    def fetchall(self):
        if self.last == 'SHOW COLUMNS FROM users':
            names = ['username', 'is_active', 'session_version', 'login_deleted_at'] if self.existing else ['username']
            return [{'Field': name} for name in names]
        return [{'Field': 'action'}]


class AccountMigrationTests(unittest.TestCase):
    def test_existing_schema_does_not_repeat_column_additions(self):
        cursor = MigrationCursor(existing=True)
        ensure_account_lifecycle_schema(cursor)
        self.assertFalse(any(sql.startswith('ALTER TABLE') for sql, _ in cursor.sql))

    def test_schema_additions_default_existing_accounts_to_active_and_restore_sql_mode(self):
        cursor = MigrationCursor()
        ensure_account_lifecycle_schema(cursor)
        sql = '\n'.join(statement for statement, _ in cursor.sql)
        self.assertIn('is_active TINYINT(1) NOT NULL DEFAULT 1', sql)
        self.assertIn('session_version INT NOT NULL DEFAULT 0', sql)
        self.assertIn('login_deleted_at DATETIME NULL DEFAULT NULL', sql)
        self.assertEqual(cursor.sql[-1], ('SET SESSION sql_mode=%s', ('STRICT_TRANS_TABLES,NO_ZERO_DATE,NO_ZERO_IN_DATE',)))
        self.assertNotIn('DELETE ', sql)

    def test_schema_failure_still_restores_sql_mode(self):
        cursor = MigrationCursor(fail=True)
        with self.assertRaises(RuntimeError): ensure_account_lifecycle_schema(cursor)
        self.assertEqual(cursor.sql[-1][0], 'SET SESSION sql_mode=%s')
        self.assertIn('NO_ZERO_DATE', cursor.sql[-1][1][0])


if __name__ == '__main__':
    unittest.main()
