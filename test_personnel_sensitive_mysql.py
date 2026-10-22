"""Opt-in NIR HTTP/storage checks against connection-scoped MySQL TEMPORARY tables.

No production requests, real employee edits, database creation, or permanent tables.
The fixture shadows the register and users tables only on its owned connection.
"""
from contextlib import ExitStack, closing
import json
import os
from pathlib import Path
import re
import tempfile
import unittest
from unittest.mock import patch

from fastapi import FastAPI, Header, HTTPException
from fastapi.testclient import TestClient
import pymysql

from backend.personnel_register import register_personnel_routes
from backend.personnel_sensitive import ensure_local_nir_key_file


TABLES = {'users', 'personnel_contracts', 'personnel_register', 'personnel_register_periods', 'personnel_register_events'}
# Fictional parts for format-only testing; these fixtures are never tied to a real person.
SYNTHETIC_NIR = '19999' + '00' + '000000' + '01'
REPLACEMENT_NIR = '29999' + '00' + '000000' + '02'


class RestrictedCursor:
    def __init__(self, fixture):
        self.fixture = fixture
        self.cursor = fixture.db.cursor()

    def __enter__(self): return self
    def __exit__(self, *_): self.cursor.close()

    @property
    def description(self): return self.cursor.description

    def execute(self, sql, params=()):
        if sql == "SHOW TABLES LIKE 'personnel_register'":
            # SHOW TABLES omits TEMPORARY tables; this fixture owns this shadow.
            return self.cursor.execute("SELECT 'personnel_register' AS table_name")
        upper = sql.lstrip().upper()
        if upper.startswith('CREATE TABLE IF NOT EXISTS'):
            name = re.match(r'\s*CREATE TABLE IF NOT EXISTS (\w+)', sql, re.I)[1]
            if name not in TABLES - {'users'}:
                raise AssertionError('Unexpected schema write refused')
            # All these names have already been created as TEMPORARY tables below.
            return 0
        if 'FROM personnel_register_periods p2' in sql:
            # MySQL cannot reopen one TEMPORARY table under both aliases of the
            # production self-subquery. Mirror only this connection's fixture rows.
            self.cursor.execute('DROP TEMPORARY TABLE IF EXISTS personnel_fixture_period_lookup')
            self.cursor.execute('CREATE TEMPORARY TABLE personnel_fixture_period_lookup AS SELECT * FROM personnel_register_periods')
            sql = sql.replace('FROM personnel_register_periods p2', 'FROM personnel_fixture_period_lookup p2')
        if not upper.startswith(('SELECT', 'SHOW')):
            mutation = re.match(r'\s*(?:UPDATE|INSERT\s+INTO|ALTER\s+TABLE)\s+(\w+)\b', sql, re.I)
            if not mutation or mutation[1] not in TABLES:
                raise AssertionError('Unexpected write refused by temporary-table fixture')
        if self.fixture.fail_audit and 'INSERT INTO personnel_register_events' in sql:
            raise RuntimeError('Injected isolated audit failure')
        serialized = repr(params)
        if SYNTHETIC_NIR in serialized or REPLACEMENT_NIR in serialized:
            raise AssertionError('A plaintext NIR reached SQL parameters')
        return self.cursor.execute(sql, params)

    def fetchone(self): return self.cursor.fetchone()
    def fetchall(self): return self.cursor.fetchall()
    @property
    def lastrowid(self): return self.cursor.lastrowid
    @property
    def rowcount(self): return self.cursor.rowcount


class TemporaryConnection:
    def __init__(self, db): self.db = db; self.fail_audit = False
    def cursor(self): return RestrictedCursor(self)
    def commit(self): self.db.commit()
    def rollback(self): self.db.rollback()
    def close(self): pass  # Only ExitStack closes the real temporary-table connection.


@unittest.skipUnless(os.getenv('PRESENCE_MYSQL_TEMP_TEST') == '1', 'opt-in local MySQL temporary-table integration')
class PersonnelSensitiveMySQLTests(unittest.TestCase):
    def setUp(self):
        self.stack = ExitStack()
        self.addCleanup(self.stack.close)
        config = json.loads((Path(__file__).parent / '.local-test/rebuild/runtime.json').read_text(encoding='utf-8-sig'))
        self.assertEqual(config['database'], 'presence_local_rebuild')
        self.db = self.stack.enter_context(closing(pymysql.connect(
            host='127.0.0.1', port=3307, user='root', password=config['db_password'],
            database=config['database'], cursorclass=pymysql.cursors.DictCursor, autocommit=False,
        )))
        with self.db.cursor() as cursor:
            cursor.execute('''CREATE TEMPORARY TABLE users (
                username VARCHAR(255) PRIMARY KEY,email VARCHAR(255),role_tags TEXT,is_active TINYINT DEFAULT 1,
                login_deleted_at DATETIME NULL,session_version INT DEFAULT 0,employee_type VARCHAR(20) DEFAULT 'salarie',
                contract_start_date DATE NULL,contract_end_date DATE NULL,is_cdi BOOLEAN DEFAULT 0) ENGINE=InnoDB''')
            cursor.execute('''CREATE TEMPORARY TABLE personnel_contracts (
                id INT AUTO_INCREMENT PRIMARY KEY,username VARCHAR(255) NOT NULL,
                employee_type VARCHAR(20) NOT NULL,start_date DATE NOT NULL,end_date DATE NULL,
                is_cdi BOOLEAN DEFAULT 0,is_archived TINYINT DEFAULT 0) ENGINE=InnoDB''')
            # Start with the old register schema to exercise the actual additive ALTER.
            cursor.execute('''CREATE TEMPORARY TABLE personnel_register (
                employee_id BIGINT AUTO_INCREMENT PRIMARY KEY,full_name VARCHAR(200) NOT NULL,
                nationality VARCHAR(100) DEFAULT '',sex VARCHAR(1) DEFAULT '',address VARCHAR(1000) DEFAULT '',
                birth_date DATE NULL,job_title VARCHAR(200) DEFAULT '',qualification VARCHAR(200) DEFAULT '',
                personnel_type VARCHAR(20) DEFAULT '',linked_username VARCHAR(255) NULL UNIQUE,
                is_archived TINYINT DEFAULT 0,version INT DEFAULT 1,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP) ENGINE=InnoDB''')
            cursor.execute('''CREATE TEMPORARY TABLE personnel_register_periods (
                period_id BIGINT AUTO_INCREMENT PRIMARY KEY,employee_id BIGINT NOT NULL,
                entry_date DATE NULL,exit_date DATE NULL,contract_type VARCHAR(20) DEFAULT '',
                other_contract VARCHAR(200) DEFAULT '',created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP) ENGINE=InnoDB''')
            cursor.execute('''CREATE TEMPORARY TABLE personnel_register_events (
                event_id BIGINT AUTO_INCREMENT PRIMARY KEY,employee_id BIGINT NOT NULL,
                actor_username VARCHAR(255),action VARCHAR(40),before_snapshot LONGTEXT NULL,
                after_snapshot LONGTEXT NOT NULL,created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP) ENGINE=InnoDB''')
            for username, role in [('fixture-admin', 'Admin'), ('fixture-manager', 'Responsable'), ('fixture-employee', 'Employe')]:
                cursor.execute('INSERT INTO users(username,email,role_tags) VALUES(%s,%s,%s)',
                               (username, username + '@example.test', json.dumps([role])))
        self.db.commit()
        self.proxy = TemporaryConnection(self.db)
        key_directory = self.stack.enter_context(tempfile.TemporaryDirectory())
        self.key_file = Path(key_directory) / 'nir-keys.json'
        ensure_local_nir_key_file(self.key_file)
        self.stack.enter_context(patch.dict(os.environ, {
            'PRESENCE_PERSONNEL_NIR_ENABLED': '1', 'PRESENCE_PERSONNEL_NIR_KEYS_FILE': str(self.key_file),
        }))

        def load_user(cursor, username):
            cursor.execute('SELECT * FROM users WHERE username=%s', (username,))
            row = cursor.fetchone()
            if row:
                row['role_tags'] = json.loads(row['role_tags'])
            return row

        def actor(authorization: str | None = Header(default=None)):
            if not authorization or not authorization.startswith('Bearer fixture-'):
                raise HTTPException(401, 'Fixture authentication required')
            # Deliberately untrusted claimed role: routes must recheck the users table.
            return {'username': authorization.removeprefix('Bearer '), 'session_version': 0, 'role_tags': ['Admin']}

        app = FastAPI()
        register_personnel_routes(app, get_db_connection=lambda: self.proxy, get_current_user=actor, load_user=load_user)
        self.client = self.stack.enter_context(TestClient(app, raise_server_exceptions=False))
        self.headers = {'Authorization': 'Bearer fixture-admin'}

    def request(self, method, path, **kwargs):
        response = self.client.request(method, '/personnel-register' + path, headers=self.headers, **kwargs)
        self.assertIn('no-store', response.headers.get('cache-control', ''))
        return response

    def create(self, number=SYNTHETIC_NIR):
        response = self.request('POST', '', json={'full_name': 'Isolated synthetic fixture', 'social_security_number': number})
        self.assertEqual(response.status_code, 201)
        self.assertNotIn(SYNTHETIC_NIR, response.text)
        return response.json()['employee']

    def stored(self, employee_id):
        with self.db.cursor() as cursor:
            cursor.execute('SELECT * FROM personnel_register WHERE employee_id=%s', (employee_id,))
            return cursor.fetchone()

    def test_encrypted_round_trip_and_private_api_responses(self):
        employee = self.create()
        identity = employee['employee_id']
        self.assertTrue(employee['has_social_security_number'])
        self.assertNotIn('social_security_number_encrypted', employee)
        ciphertext = self.stored(identity)['social_security_number_encrypted']
        self.assertTrue(ciphertext)
        self.assertNotIn(SYNTHETIC_NIR, ciphertext)
        for path in ['', f'/{identity}', f'/{identity}/history']:
            response = self.request('GET', path)
            self.assertEqual(response.status_code, 200)
            self.assertNotIn(SYNTHETIC_NIR, response.text)
            self.assertNotIn(ciphertext, response.text)
        reveal = self.request('POST', f'/{identity}/social-security-number/reveal')
        self.assertEqual(reveal.status_code, 200)
        self.assertEqual(reveal.json()['social_security_number'], SYNTHETIC_NIR)
        history = self.request('GET', f'/{identity}/history').json()['events']
        self.assertIn('social_security_number_viewed', [event['action'] for event in history])
        with self.db.cursor() as cursor:
            cursor.execute('SELECT before_snapshot,after_snapshot FROM personnel_register_events')
            snapshots = json.dumps(cursor.fetchall())
        self.assertNotIn(SYNTHETIC_NIR, snapshots)
        self.assertNotIn(ciphertext, snapshots)

    def test_omission_preserves_replacement_and_clear_are_versioned(self):
        employee = self.create()
        identity = employee['employee_id']
        ciphertext = self.stored(identity)['social_security_number_encrypted']
        changed = self.request('PUT', f'/{identity}', json={'full_name': 'Renamed fixture', 'expected_version': employee['version']})
        self.assertEqual(changed.status_code, 200)
        employee = changed.json()['employee']
        self.assertEqual(self.stored(identity)['social_security_number_encrypted'], ciphertext)
        replace = self.request('PUT', f'/{identity}', json={'full_name': 'Renamed fixture', 'social_security_number': REPLACEMENT_NIR, 'expected_version': employee['version']})
        self.assertEqual(replace.status_code, 200)
        employee = replace.json()['employee']
        self.assertNotIn(REPLACEMENT_NIR, replace.text)
        self.assertGreater(employee['version'], 1)
        stale = self.request('PUT', f'/{identity}', json={'full_name': 'Stale fixture', 'social_security_number': None, 'expected_version': 1})
        self.assertEqual(stale.status_code, 409)
        cleared = self.request('PUT', f'/{identity}', json={'full_name': 'Renamed fixture', 'social_security_number': None, 'expected_version': employee['version']})
        self.assertEqual(cleared.status_code, 200)
        self.assertFalse(cleared.json()['employee']['has_social_security_number'])
        self.assertIsNone(self.stored(identity)['social_security_number_encrypted'])

    def test_denied_roles_and_disabled_admin_cannot_reveal(self):
        identity = self.create()['employee_id']
        for actor in ['fixture-manager', 'fixture-employee']:
            response = self.client.post(f'/personnel-register/{identity}/social-security-number/reveal', headers={'Authorization': 'Bearer ' + actor})
            self.assertEqual(response.status_code, 403)
            self.assertNotIn(SYNTHETIC_NIR, response.text)
        self.assertEqual(self.client.post(f'/personnel-register/{identity}/social-security-number/reveal').status_code, 401)
        with self.db.cursor() as cursor:
            cursor.execute("UPDATE users SET is_active=0 WHERE username='fixture-admin'")
        self.db.commit()
        self.assertEqual(self.request('POST', f'/{identity}/social-security-number/reveal').status_code, 401)

    def test_missing_key_and_audit_failure_do_not_release_plaintext(self):
        identity = self.create()['employee_id']
        self.proxy.fail_audit = True
        response = self.request('POST', f'/{identity}/social-security-number/reveal')
        self.assertGreaterEqual(response.status_code, 500)
        self.assertNotIn(SYNTHETIC_NIR, response.text)
        self.proxy.fail_audit = False
        with patch.dict(os.environ, {'PRESENCE_PERSONNEL_NIR_KEYS_FILE': str(self.key_file.with_name('absent.json'))}):
            response = self.request('POST', f'/{identity}/social-security-number/reveal')
            self.assertEqual(response.status_code, 503)
            self.assertNotIn(SYNTHETIC_NIR, response.text)
            self.assertEqual(self.request('GET', f'/{identity}').status_code, 200)
            self.assertFalse(self.request('GET', '/capabilities').json()['nir_enabled'])

    def test_invalid_input_and_feature_switch_fail_closed(self):
        marker = 'not-a-valid-synthetic-nir'
        response = self.request('POST', '', json={'full_name': 'Invalid fixture', 'social_security_number': marker})
        self.assertEqual(response.status_code, 422)
        self.assertNotIn(marker, response.text)
        with patch.dict(os.environ, {'PRESENCE_PERSONNEL_NIR_ENABLED': '0'}):
            response = self.request('POST', '', json={'full_name': 'Disabled fixture', 'social_security_number': SYNTHETIC_NIR})
            self.assertEqual(response.status_code, 503)
            ordinary = self.request('POST', '', json={'full_name': 'Non-sensitive fixture'})
            self.assertEqual(ordinary.status_code, 201)


if __name__ == '__main__':
    unittest.main()
