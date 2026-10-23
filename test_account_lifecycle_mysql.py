"""Opt-in MySQL check using connection-scoped TEMPORARY tables only.

Requires the existing isolated local MySQL on 127.0.0.1:3307. Never imports
local_rebuild, starts migrations, creates a database, or updates real user rows.
"""
from contextlib import closing
import json
import os
from pathlib import Path
import re
import secrets
import unittest
from unittest.mock import patch

import pymysql
import test_backend_helpers
from backend import main
from backend.account_lifecycle import register_account_lifecycle
from fastapi import Depends, FastAPI
from fastapi.testclient import TestClient


class RestrictedCursor:
    def __init__(self, connection):
        self.connection = connection
        self.cursor = connection.db.cursor()
    def __enter__(self): return self
    def __exit__(self, *_): self.cursor.close()

    @property
    def description(self): return self.cursor.description

    def execute(self, sql, params=()):
        if sql == "SHOW TABLES LIKE 'personnel_register'":
            # This fixture has no register; do not read any real personnel table.
            return self.cursor.execute('SELECT NULL WHERE 1=0')
        if self.connection.fail_event and 'INSERT INTO account_status_events' in sql:
            raise RuntimeError('Injected isolated audit failure')
        if not sql.lstrip().upper().startswith('SELECT'):
            mutation = re.match(r'\s*(?:UPDATE|INSERT\s+INTO)\s+(\w+)\b', sql, re.I)
            if not mutation or mutation[1] not in {'users', 'account_status_events'}:
                raise AssertionError('Unexpected write refused by isolated MySQL fixture')
        return self.cursor.execute(sql, params)

    def fetchone(self): return self.cursor.fetchone()
    def fetchall(self): return self.cursor.fetchall()


class TemporaryConnection:
    def __init__(self, db): self.db = db; self.fail_event = False
    def cursor(self): return RestrictedCursor(self)
    def commit(self): self.db.commit()
    def rollback(self): self.db.rollback()
    def close(self): pass  # The test owns and closes the shared temporary-table connection.


@unittest.skipUnless(os.getenv('PRESENCE_MYSQL_TEMP_TEST') == '1', 'opt-in local MySQL TEMPORARY-table integration')
class AccountLifecycleMySQLTests(unittest.TestCase):
    def test_real_mysql_retention_authentication_and_two_step_delete(self):
        config = json.loads((Path(__file__).parent / '.local-test/rebuild/runtime.json').read_text(encoding='utf-8-sig'))
        self.assertEqual(config['database'], 'presence_local_rebuild')
        with closing(pymysql.connect(host='127.0.0.1', port=3307, user='root', password=config['db_password'],
                                     database=config['database'], cursorclass=pymysql.cursors.DictCursor,
                                     autocommit=False)) as db:
            with db.cursor() as cursor:
                # Shadow the real tables for this connection; other app connections cannot see them.
                cursor.execute('''CREATE TEMPORARY TABLE users (
                    username VARCHAR(255) PRIMARY KEY,email VARCHAR(255) UNIQUE,role_tags TEXT,password_hash TEXT,
                    is_active TINYINT NOT NULL DEFAULT 1,session_version INT NOT NULL DEFAULT 0,
                    login_failed_attempts INT NOT NULL DEFAULT 0,login_locked_until DATETIME(6) NULL,login_last_failed_at DATETIME(6) NULL,
                    login_deleted_at DATETIME NULL,remember_token_hash CHAR(64) NULL,remember_token_expires DATETIME NULL,
                    reset_token TEXT NULL,reset_token_expires DATETIME NULL) ENGINE=InnoDB''')
                cursor.execute('''CREATE TEMPORARY TABLE account_status_events (
                    id BIGINT AUTO_INCREMENT PRIMARY KEY,target_username VARCHAR(255),actor_username VARCHAR(255),
                    was_active TINYINT,is_active TINYINT,action VARCHAR(20) DEFAULT 'activation',
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP) ENGINE=InnoDB''')
                cursor.execute('''CREATE TEMPORARY TABLE lifecycle_retained_business (
                    kind VARCHAR(40),owner VARCHAR(255),uploaded_by VARCHAR(255),archived_by VARCHAR(255),payload TEXT) ENGINE=InnoDB''')
                secret = secrets.token_urlsafe(24)
                hashed = main.hash_password(secret)
                for username, role in [('fixture-admin', 'SuperAdmin'), ('fixture-employee', 'Employe')]:
                    cursor.execute('INSERT INTO users(username,email,role_tags,password_hash) VALUES(%s,%s,%s,%s)',
                                   (username, username + '@example.test', json.dumps([role]), hashed))
                for kind, owner in [('attendance', 'fixture-employee'), ('contract', 'fixture-employee'),
                                    ('document', 'another-person'), ('leave', 'fixture-employee')]:
                    cursor.execute('INSERT INTO lifecycle_retained_business VALUES(%s,%s,%s,%s,%s)',
                                   (kind, owner, 'fixture-employee', 'fixture-employee', 'unchanged fixture'))
            db.commit()
            proxy = TemporaryConnection(db)

            def load_user(cursor, username):
                cursor.execute('SELECT * FROM users WHERE username=%s', (username,))
                row = cursor.fetchone()
                if row:
                    tags = json.loads(row['role_tags'])
                    row.update(role_tags=tags, role=tags[0], is_admin='Admin' in tags)
                return row

            def read_business():
                with db.cursor() as cursor:
                    cursor.execute('SELECT * FROM lifecycle_retained_business ORDER BY kind')
                    return cursor.fetchall()

            app = FastAPI()
            register_account_lifecycle(app, get_db_connection=lambda: proxy, get_current_user=main.get_current_user,
                                       load_user=load_user, public_user=main.public_user, hash_password=main.hash_password)
            app.add_api_route('/login', main.login, methods=['POST'])
            app.add_api_route('/restore-session', main.restore_session, methods=['POST'])

            @app.get('/private')
            def private(actor=Depends(main.get_current_user)): return {'username': actor['username']}

            with patch.object(main, 'get_db_connection', lambda: proxy), patch.object(main, 'load_user', load_user), TestClient(app, raise_server_exceptions=False) as client:
                admin = {'Authorization': 'Bearer ' + main.create_session_token('fixture-admin')}
                login = lambda: client.post('/login', json={'username': 'fixture-employee', 'password': secret, 'remember_me': True})
                session = login().json()
                by_email = client.post('/login', json={'username': '  FIXTURE-EMPLOYEE@EXAMPLE.TEST  ', 'password': secret})
                self.assertEqual(by_email.status_code, 200)
                self.assertEqual(main.read_session_token(by_email.json()['auth_token']), 'fixture-employee')
                old = {'Authorization': 'Bearer ' + session['auth_token']}
                target = '/users/fixture-employee'
                before = read_business()
                delete = lambda: client.request('DELETE', target + '/login', headers=admin, json={'confirmation': 'fixture-employee'})
                activate = lambda value: client.patch(target + '/activation', headers=admin, json={'is_active': value})
                self.assertEqual(delete().status_code, 409)
                proxy.fail_event = True
                self.assertEqual(activate(False).status_code, 500)
                self.assertEqual(client.get('/private', headers=old).status_code, 200)
                proxy.fail_event = False
                self.assertEqual(activate(False).status_code, 200)
                self.assertEqual(client.get('/private', headers=old).status_code, 401)
                self.assertEqual(login().status_code, 401)
                self.assertEqual(activate(True).status_code, 200)
                self.assertEqual(client.get('/private', headers=old).status_code, 401)
                self.assertEqual(login().status_code, 200)
                self.assertEqual(activate(False).status_code, 200)
                proxy.fail_event = True
                self.assertEqual(delete().status_code, 500)
                with proxy.cursor() as cursor:
                    row = load_user(cursor, 'fixture-employee')
                self.assertIsNone(row['login_deleted_at'])
                self.assertEqual(row['password_hash'], hashed)
                proxy.fail_event = False
                self.assertEqual(delete().status_code, 200)
                self.assertEqual(delete().status_code, 200)
                self.assertEqual(activate(True).status_code, 409)
                self.assertEqual(login().status_code, 401)
                self.assertEqual(client.post('/restore-session', json={'remember_token': session['remember_token']}).status_code, 401)
                self.assertEqual(read_business(), before)
                events = client.get(target + '/activation-history', headers=admin).json()['events']
                self.assertEqual(len(events), 4)
                self.assertEqual(events[0]['action'], 'login_deleted')
            # Closing this connection drops every temporary fixture table, including on failure.


if __name__ == '__main__':
    unittest.main()
