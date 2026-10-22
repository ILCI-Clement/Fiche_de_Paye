"""Isolated personnel HTTP/storage checks using temporary SQLite and synthetic values.

No production services, runtime credentials, real people or persistent business tables.
SQLite adapts SQL syntax only; real MySQL schema/locking is checked separately.
"""
from contextlib import closing
from datetime import date, datetime, timezone
import json
import os
from pathlib import Path
import sqlite3
import tempfile
import unittest
from unittest.mock import patch

from fastapi import FastAPI, HTTPException
from fastapi.testclient import TestClient

from backend.personnel_register import register_personnel_routes
from backend.personnel_sensitive import ensure_local_nir_key_file
from backend.personnel_store import ensure_personnel_schema, safe_audit_snapshot


# Fictional format fixtures: month 99 cannot identify an actual birth date.
SYNTHETIC_NIR = '19999' + '00' + '000000' + '01'
REPLACEMENT_NIR = '29999' + '00' + '000000' + '02'


class Cursor:
    def __init__(self, connection):
        self.connection = connection
        self.cursor = connection.db.cursor()

    def __enter__(self): return self
    def __exit__(self, *_): self.cursor.close()

    def execute(self, sql, params=()):
        # A dedicated number must never reach SQL, including snapshots and audit events.
        if SYNTHETIC_NIR in repr(params) or REPLACEMENT_NIR in repr(params):
            raise AssertionError('Plaintext sensitive fixture reached SQL')
        if self.connection.fixture.fail_audit and 'INSERT INTO personnel_register_events' in sql:
            raise RuntimeError('Injected isolated audit failure')
        if sql == "SHOW TABLES LIKE 'personnel_register'":
            return self.cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='personnel_register'")
        if 'FOR UPDATE' in sql:
            if not self.connection.db.in_transaction:
                self.connection.db.execute('BEGIN IMMEDIATE')
            sql = sql.replace(' FOR UPDATE', '')
        return self.cursor.execute(sql.replace('%s', '?'), params)

    @staticmethod
    def record(row):
        if row is None:
            return None
        result = dict(row)
        for field in ('birth_date', 'entry_date', 'exit_date', 'period_entry_date', 'period_exit_date'):
            if result.get(field):
                result[field] = date.fromisoformat(result[field])
        return result

    def fetchone(self): return self.record(self.cursor.fetchone())
    def fetchall(self): return [self.record(row) for row in self.cursor.fetchall()]
    @property
    def lastrowid(self): return self.cursor.lastrowid
    @property
    def description(self): return self.cursor.description


class Connection:
    def __init__(self, fixture):
        self.fixture = fixture
        self.db = sqlite3.connect(fixture.path, timeout=10, isolation_level=None)
        self.db.row_factory = sqlite3.Row
        self.db.create_function('UNIX_TIMESTAMP', 1, lambda value:
            datetime.fromisoformat(value).replace(tzinfo=timezone.utc).timestamp() if value else None)

    def cursor(self): return Cursor(self)
    def commit(self): self.db.commit()
    def rollback(self): self.db.rollback()
    def close(self): self.db.close()


class PersonnelSensitiveHTTPTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.path = Path(self.temp.name) / 'isolated.sqlite3'
        self.key_file = Path(self.temp.name) / 'independent-keys.json'
        ensure_local_nir_key_file(self.key_file)
        self.env = patch.dict(os.environ, {
            'PRESENCE_PERSONNEL_NIR_ENABLED': '1',
            'PRESENCE_PERSONNEL_NIR_KEYS_FILE': str(self.key_file),
        })
        self.env.start()
        self.addCleanup(self.env.stop)
        self.fail_audit = False
        with closing(sqlite3.connect(self.path)) as db, db:
            db.executescript('''CREATE TABLE users (
                username TEXT PRIMARY KEY,email TEXT UNIQUE,role_tags TEXT,
                is_active INTEGER NOT NULL DEFAULT 1,session_version INTEGER NOT NULL DEFAULT 0,
                login_deleted_at TEXT NULL);
                CREATE TABLE personnel_register (
                employee_id INTEGER PRIMARY KEY AUTOINCREMENT,full_name TEXT NOT NULL,
                nationality TEXT NOT NULL DEFAULT '',sex TEXT NOT NULL DEFAULT '',
                address TEXT NOT NULL DEFAULT '',birth_date TEXT NULL,
                job_title TEXT NOT NULL DEFAULT '',qualification TEXT NOT NULL DEFAULT '',
                personnel_type TEXT NOT NULL DEFAULT '',linked_username TEXT NULL UNIQUE,
                social_security_number_encrypted TEXT NULL,is_archived INTEGER NOT NULL DEFAULT 0,
                version INTEGER NOT NULL DEFAULT 1,created_at TEXT DEFAULT CURRENT_TIMESTAMP,
                updated_at TEXT DEFAULT CURRENT_TIMESTAMP);
                CREATE TABLE personnel_register_periods (
                period_id INTEGER PRIMARY KEY AUTOINCREMENT,employee_id INTEGER NOT NULL,
                entry_date TEXT NULL,exit_date TEXT NULL,contract_type TEXT NOT NULL DEFAULT '',
                other_contract TEXT NOT NULL DEFAULT '',created_at TEXT DEFAULT CURRENT_TIMESTAMP,
                updated_at TEXT DEFAULT CURRENT_TIMESTAMP);
                CREATE TABLE personnel_register_events (
                event_id INTEGER PRIMARY KEY AUTOINCREMENT,employee_id INTEGER NOT NULL,
                actor_username TEXT NOT NULL,action TEXT NOT NULL,before_snapshot TEXT NULL,
                after_snapshot TEXT NOT NULL,created_at TEXT DEFAULT CURRENT_TIMESTAMP);''')
            for username, role in [('fixture-admin', 'Admin'), ('fixture-manager', 'Responsable'),
                                   ('fixture-employee', 'Employe')]:
                db.execute('INSERT INTO users(username,email,role_tags) VALUES(?,?,?)',
                           (username, username + '@example.test', json.dumps([role])))
        self.actor = {'username': 'fixture-admin', 'role_tags': ['Admin'], 'session_version': 0}

        def authenticated():
            if self.actor is None:
                raise HTTPException(401, 'Authentication required')
            return dict(self.actor)

        def load_user(cursor, username):
            cursor.execute('SELECT * FROM users WHERE username=%s', (username,))
            row = cursor.fetchone()
            if row:
                row['role_tags'] = json.loads(row['role_tags'])
            return row

        app = FastAPI()
        register_personnel_routes(app, get_db_connection=lambda: Connection(self),
                                  get_current_user=authenticated, load_user=load_user)
        # Do not enter TestClient's lifespan: the fixture has its own isolated schema.
        self.client = TestClient(app, raise_server_exceptions=False)
        self.addCleanup(self.client.close)

    def create(self, **values):
        return self.client.post('/personnel-register', json={'full_name': 'Synthetic personnel fixture', **values})

    def create_number(self):
        response = self.create(social_security_number=SYNTHETIC_NIR)
        self.assertEqual(response.status_code, 201, response.text)
        return response.json()['employee']

    def update(self, employee, **values):
        return self.client.put('/personnel-register/' + str(employee['employee_id']), json={
            'full_name': employee['full_name'], 'expected_version': employee['version'], **values,
        })

    def reveal(self, employee):
        return self.client.post(f"/personnel-register/{employee['employee_id']}/social-security-number/reveal")

    def rows(self, table):
        self.assertIn(table, {'users', 'personnel_register', 'personnel_register_periods', 'personnel_register_events'})
        with closing(sqlite3.connect(self.path)) as db:
            db.row_factory = sqlite3.Row
            return [dict(row) for row in db.execute('SELECT * FROM ' + table)]

    def stored(self, employee):
        return next(row for row in self.rows('personnel_register') if row['employee_id'] == employee['employee_id'])

    def raw_update(self, sql, params):
        with closing(sqlite3.connect(self.path)) as db, db:
            db.execute(sql, params)

    def assert_safe(self, value):
        serialized = value if isinstance(value, str) else json.dumps(value, default=str)
        for excluded in (SYNTHETIC_NIR, REPLACEMENT_NIR, 'social_security_number_encrypted',
                         'ciphertext', 'nonce', 'algorithm'):
            self.assertNotIn(excluded, serialized)

    def assert_no_cache(self, response):
        self.assertIn('no-store', response.headers.get('Cache-Control', ''))
        self.assertIn('no-cache', response.headers.get('Cache-Control', ''))
        self.assertEqual(response.headers.get('Pragma'), 'no-cache')

    def test_create_encrypts_and_reveal_commits_metadata_audit(self):
        employee = self.create_number()
        self.assertTrue(employee['has_social_security_number'])
        self.assert_safe(employee)
        stored = self.stored(employee)['social_security_number_encrypted']
        self.assertNotIn(SYNTHETIC_NIR, stored)
        self.assertEqual(json.loads(stored)['algorithm'], 'A256GCM')
        response = self.reveal(employee)
        self.assertEqual(response.status_code, 200, response.text)
        self.assertEqual(response.json(), {'social_security_number': SYNTHETIC_NIR})
        self.assert_no_cache(response)
        events = self.rows('personnel_register_events')
        self.assertEqual([row['action'] for row in events],
                         ['registered', 'social_security_number_changed', 'social_security_number_viewed'])
        self.assert_safe(events)
        for row in events[1:]:
            after = json.loads(row['after_snapshot'])
            self.assertEqual(set(after['employee']), {'employee_id', 'has_social_security_number'})

    def test_list_detail_and_history_never_decrypt_or_expose_ciphertext(self):
        employee = self.create_number()
        with patch('backend.personnel_register.decrypt_social_security_number', side_effect=AssertionError('Unexpected decrypt')):
            for suffix in ('', f"/{employee['employee_id']}", f"/{employee['employee_id']}/history"):
                response = self.client.get('/personnel-register' + suffix)
                self.assertEqual(response.status_code, 200, response.text)
                self.assert_safe(response.json())
                self.assert_no_cache(response)

    def test_omitted_number_preserves_envelope_during_ordinary_edit(self):
        employee = self.create_number()
        envelope = self.stored(employee)['social_security_number_encrypted']
        response = self.update(employee, job_title='Synthetic job')
        self.assertEqual(response.status_code, 200, response.text)
        self.assertTrue(response.json()['employee']['has_social_security_number'])
        self.assertEqual(self.stored(employee)['social_security_number_encrypted'], envelope)

    def test_replacement_and_explicit_null_clear_have_safe_change_events(self):
        employee = self.create_number()
        response = self.update(employee, social_security_number=REPLACEMENT_NIR)
        self.assertEqual(response.status_code, 200, response.text)
        employee = response.json()['employee']
        self.assertEqual(employee['version'], 2)
        self.assertEqual(self.reveal(employee).json()['social_security_number'], REPLACEMENT_NIR)
        response = self.update(employee, social_security_number=None)
        self.assertEqual(response.status_code, 200, response.text)
        employee = response.json()['employee']
        self.assertFalse(employee['has_social_security_number'])
        self.assertEqual(employee['version'], 3)
        self.assertIsNone(self.stored(employee)['social_security_number_encrypted'])
        self.assertEqual(self.reveal(employee).status_code, 404)
        events = self.rows('personnel_register_events')
        self.assertEqual(sum(row['action'] == 'social_security_number_changed' for row in events), 3)
        self.assert_safe(events)

    def test_blank_number_clears_and_already_empty_clear_is_a_noop(self):
        employee = self.create_number()
        response = self.update(employee, social_security_number='   ')
        self.assertEqual(response.status_code, 200, response.text)
        employee = response.json()['employee']
        count = len(self.rows('personnel_register_events'))
        again = self.update(employee, social_security_number=None)
        self.assertEqual(again.status_code, 200, again.text)
        self.assertEqual(again.json()['employee']['version'], employee['version'])
        self.assertEqual(len(self.rows('personnel_register_events')), count)

    def test_capability_requires_exact_flag_and_available_independent_key(self):
        self.assertEqual(self.client.get('/personnel-register/capabilities').json(), {'nir_enabled': True})
        for flag in ('0', '', 'true'):
            with patch.dict(os.environ, {'PRESENCE_PERSONNEL_NIR_ENABLED': flag}):
                response = self.client.get('/personnel-register/capabilities')
                self.assertEqual(response.json(), {'nir_enabled': False})
                self.assert_no_cache(response)
        with patch.dict(os.environ, {'PRESENCE_PERSONNEL_NIR_KEYS_FILE': str(self.key_file) + '.missing'}):
            self.assertEqual(self.client.get('/personnel-register/capabilities').json(), {'nir_enabled': False})

    def test_disabled_or_missing_key_blocks_only_dedicated_number_actions(self):
        employee = self.create_number()
        initial = self.stored(employee)['social_security_number_encrypted']
        for configuration in ({'PRESENCE_PERSONNEL_NIR_ENABLED': '0'},
                              {'PRESENCE_PERSONNEL_NIR_KEYS_FILE': str(self.key_file) + '.missing'}):
            with patch.dict(os.environ, configuration):
                for response in (self.create(social_security_number=SYNTHETIC_NIR), self.reveal(employee),
                                 self.update(employee, social_security_number=None)):
                    self.assertEqual(response.status_code, 503, response.text)
                    self.assert_safe(response.text)
                    self.assert_no_cache(response)
                ordinary = self.create()
                self.assertEqual(ordinary.status_code, 201, ordinary.text)
                updated = self.update(employee, job_title='Ordinary edit without keys')
                self.assertEqual(updated.status_code, 200, updated.text)
                employee = updated.json()['employee']
                self.assertEqual(self.stored(employee)['social_security_number_encrypted'], initial)

    def test_numberless_create_and_explicit_null_do_not_invent_a_number(self):
        response = self.create(social_security_number=None)
        self.assertEqual(response.status_code, 201, response.text)
        self.assertFalse(response.json()['employee']['has_social_security_number'])
        self.assertEqual([row['action'] for row in self.rows('personnel_register_events')], ['registered'])

    def test_bad_number_and_unknown_fields_return_redacted_validation_errors(self):
        for payload in ({'full_name': 'Synthetic', 'social_security_number': SYNTHETIC_NIR + 'X'},
                        {'full_name': 'Synthetic', 'social_security_number': {'value': SYNTHETIC_NIR}},
                        {'full_name': 'Synthetic', 'nir': SYNTHETIC_NIR},
                        {'full_name': 'Synthetic', SYNTHETIC_NIR: 'unexpected key'},
                        {'full_name': '', 'social_security_number': SYNTHETIC_NIR}):
            response = self.client.post('/personnel-register', json=payload)
            self.assertEqual(response.status_code, 422, response.text)
            self.assert_safe(response.text)
            self.assertNotIn('input', response.json()['detail'][0])
            self.assert_no_cache(response)
        self.assertEqual(self.rows('personnel_register'), [])

    def test_current_roles_are_enforced_for_all_number_endpoints(self):
        employee = self.create_number()
        for username, role in [('fixture-manager', 'Responsable'), ('fixture-employee', 'Employe')]:
            self.actor = {'username': username, 'role_tags': [role], 'session_version': 0}
            for response in (self.client.get('/personnel-register/capabilities'), self.reveal(employee),
                             self.client.get('/personnel-register'),
                             self.update(employee, social_security_number=REPLACEMENT_NIR),
                             self.create(social_security_number=SYNTHETIC_NIR)):
                self.assertEqual(response.status_code, 403, response.text)
                self.assert_no_cache(response)
        self.assertEqual(len(self.rows('personnel_register')), 1)
        self.assert_safe(self.rows('personnel_register_events'))

    def test_missing_or_revoked_session_does_not_reveal(self):
        employee = self.create_number()
        for field, value in [('is_active', 0), ('role_tags', '["Employe"]'),
                             ('session_version', 1), ('login_deleted_at', '2026-10-22')]:
            self.raw_update(f'UPDATE users SET {field}=? WHERE username=?', (value, 'fixture-admin'))
            self.assertEqual(self.reveal(employee).status_code, 403 if field == 'role_tags' else 401)
            self.raw_update('UPDATE users SET is_active=1,role_tags=?,session_version=0,login_deleted_at=NULL WHERE username=?',
                            ('["Admin"]', 'fixture-admin'))
        self.actor = None
        response = self.reveal(employee)
        self.assertEqual(response.status_code, 401)
        self.assert_no_cache(response)

    def test_stale_version_cannot_replace_number(self):
        employee = self.create_number()
        self.assertEqual(self.update(employee, job_title='Another version').status_code, 200)
        response = self.update(employee, social_security_number=REPLACEMENT_NIR)
        self.assertEqual(response.status_code, 409)
        self.assertEqual(self.reveal(employee).json()['social_security_number'], SYNTHETIC_NIR)

    def test_ciphertext_tampering_and_cross_employee_copy_fail_closed(self):
        employee = self.create_number()
        other = self.create().json()['employee']
        envelope = self.stored(employee)['social_security_number_encrypted']
        self.raw_update('UPDATE personnel_register SET social_security_number_encrypted=? WHERE employee_id=?',
                        (envelope, other['employee_id']))
        response = self.reveal(other)
        self.assertEqual(response.status_code, 503, response.text)
        tampered = json.loads(envelope)
        tampered['ciphertext'] = ('A' if tampered['ciphertext'][0] != 'A' else 'B') + tampered['ciphertext'][1:]
        self.raw_update('UPDATE personnel_register SET social_security_number_encrypted=? WHERE employee_id=?',
                        (json.dumps(tampered), employee['employee_id']))
        response = self.reveal(employee)
        self.assertEqual(response.status_code, 503, response.text)
        self.assert_safe(response.text)
        self.assertFalse(any(row['action'] == 'social_security_number_viewed' for row in self.rows('personnel_register_events')))

    def test_audit_failure_rolls_back_create_and_replacement(self):
        employee = self.create_number()
        before = self.stored(employee)
        self.fail_audit = True
        with self.assertLogs('backend.personnel_register', level='ERROR') as captured:
            self.assertEqual(self.create(social_security_number=REPLACEMENT_NIR).status_code, 500)
            response = self.update(employee, social_security_number=REPLACEMENT_NIR)
            self.assertEqual(response.status_code, 500, response.text)
        self.assertEqual(self.stored(employee), before)
        self.assertEqual(len(self.rows('personnel_register')), 1)
        self.assert_safe(captured.output)

    def test_empty_ciphertext_is_corruption_not_a_cleared_number(self):
        employee = self.create_number()
        self.raw_update('UPDATE personnel_register SET social_security_number_encrypted=? WHERE employee_id=?',
                        ('', employee['employee_id']))
        detail = self.client.get(f"/personnel-register/{employee['employee_id']}")
        self.assertTrue(detail.json()['employee']['has_social_security_number'])
        response = self.reveal(employee)
        self.assertEqual(response.status_code, 503, response.text)
        self.assert_safe(response.text)
        self.assertEqual(len(self.rows('personnel_register_events')), 2)

    def test_audit_failure_never_returns_decrypted_plaintext(self):
        employee = self.create_number()
        self.fail_audit = True
        with self.assertLogs('backend.personnel_register', level='ERROR') as captured:
            response = self.reveal(employee)
        self.assertEqual(response.status_code, 500, response.text)
        self.assert_safe(response.text)
        self.assert_safe(captured.output)
        self.assert_no_cache(response)
        self.assertEqual(len(self.rows('personnel_register_events')), 2)

    def test_safety_filter_removes_dedicated_material_from_nested_audit(self):
        data = {'employee': {'has_social_security_number': True, 'social_security_number': SYNTHETIC_NIR},
                'nested': [{'social_security_number_encrypted': 'ciphertext', 'ordinary': 'retained'}]}
        self.assertEqual(safe_audit_snapshot(data), {
            'employee': {'has_social_security_number': True}, 'nested': [{'ordinary': 'retained'}],
        })


class PersonnelSensitiveMigrationTests(unittest.TestCase):
    def test_missing_encrypted_column_is_added_with_connection_scoped_show(self):
        class SchemaCursor:
            def __init__(self): self.commands = []; self.column_exists = False; self.result = None
            def execute(self, sql, params=()):
                self.commands.append(sql)
                if 'information_schema.columns' in sql:
                    self.result = {'charset_name': 'utf8mb4', 'collation_name': 'utf8mb4_unicode_ci'}
                elif sql.startswith('SHOW COLUMNS FROM personnel_register'):
                    self.result = {'Field': 'social_security_number_encrypted'} if self.column_exists else None
                elif sql.startswith('ALTER TABLE personnel_register ADD COLUMN'):
                    self.column_exists = True
                return 0
            def fetchone(self): return self.result
        cursor = SchemaCursor()
        ensure_personnel_schema(cursor)
        ensure_personnel_schema(cursor)
        additions = [sql for sql in cursor.commands if sql.startswith('ALTER TABLE personnel_register ADD COLUMN')]
        self.assertEqual(len(additions), 1)
        self.assertIn('social_security_number_encrypted TEXT NULL', additions[0])
        self.assertFalse(any(sql.lstrip().upper().startswith(('DELETE ', 'TRUNCATE ', 'DROP '))
                             for sql in cursor.commands))


if __name__ == '__main__':
    unittest.main()
