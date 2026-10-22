"""Isolated public-name and permission regressions; synthetic SQLite only.

No runtime configuration, production database, actual person, e-mail or signing
service is used. HTTP clients never enter application startup/lifespan handlers.
"""
from contextlib import closing
from datetime import datetime, timezone
import json
from pathlib import Path
import re
import sqlite3
import tempfile
import unittest
from unittest.mock import Mock

from fastapi import FastAPI, HTTPException
from fastapi.testclient import TestClient
from pymysql.err import OperationalError, ProgrammingError

import test_backend_helpers  # Synthetic import defaults; does not start the server.
from backend.person_names import (UNNAMED_PERSON, annotate_person_names, load_display_names,
                                  load_person_accounts, load_person_names, person_display_name)
from backend.person_profiles import register_profile_routes
from backend.personnel_register import register_personnel_routes
from backend.team_directory import register_team_directory
from organization_model import person_label


class PersonDisplayNameTests(unittest.TestCase):
    def test_linked_name_has_priority_and_retains_unicode(self):
        account = {'username': 'private-login', 'email': 'fallback@example.invalid',
                   'full_name': 'Stale name'}
        self.assertEqual(person_display_name(account, {'private-login': '  Élodie 张  '}), 'Élodie 张')

    def test_explicit_mapping_does_not_trust_unlinked_name(self):
        account = {'username': 'private-login', 'email': 'fallback@example.invalid',
                   'full_name': 'Unlinked name', 'display_name': 'Untrusted display'}
        self.assertEqual(person_display_name(account, {}), 'fallback')

    def test_email_local_part_without_username_fallback(self):
        self.assertEqual(person_display_name({'username': 'private-login',
                                             'email': '  prefix@example.invalid  '}), 'prefix')
        self.assertEqual(person_display_name({'username': 'private-login', 'name': 'Legacy'}), UNNAMED_PERSON)
        self.assertEqual(person_display_name(None), UNNAMED_PERSON)

    def test_empty_name_falls_back_to_email(self):
        self.assertEqual(person_display_name({'username': 'private-login', 'email': 'prefix@example.invalid'},
                                            {'private-login': ' \n '}), 'prefix')

    def test_invalid_email_cannot_create_a_display_name(self):
        for email in ('no-at', '@domain.invalid', 'local@', 'local@@domain.invalid',
                      'private login@domain.invalid', 'local@domain invalid', 'local\n@domain.invalid'):
            with self.subTest(email=email):
                self.assertEqual(person_display_name({'username': 'private-login', 'email': email}), UNNAMED_PERSON)

    def test_organization_email_rules_match_public_helper(self):
        for email in ('normal@example.invalid', '  normal@example.invalid  ', 'local@@domain.invalid',
                      'private login@domain.invalid', 'local@domain invalid', None):
            with self.subTest(email=email):
                self.assertEqual(person_label({'email': email}), person_display_name({'email': email}))
                if email not in ('normal@example.invalid', '  normal@example.invalid  '):
                    self.assertEqual(person_label({'email': email}), UNNAMED_PERSON)

    def test_same_names_keep_distinct_identity_and_authorization_metadata(self):
        accounts = [{'username': 'id-a', 'email': 'a@example.invalid', 'manager_username': 'manager-a'},
                    {'username': 'id-b', 'email': 'b@example.invalid', 'manager_username': 'manager-b'}]
        result = annotate_person_names(accounts, {'id-a': 'Same Name', 'id-b': 'Same Name'})
        self.assertEqual([row['display_name'] for row in result], ['Same Name', 'Same Name'])
        self.assertEqual([row['username'] for row in result], ['id-a', 'id-b'])
        self.assertEqual([row['manager_username'] for row in result], ['manager-a', 'manager-b'])


class PersonNameStorageTests(unittest.TestCase):
    def setUp(self):
        self.db = sqlite3.connect(':memory:')
        self.addCleanup(self.db.close)
        self.db.row_factory = sqlite3.Row
        self.db.execute('CREATE TABLE users(username TEXT PRIMARY KEY,email TEXT,password_hash TEXT)')
        self.db.executemany('INSERT INTO users VALUES(?,?,?)', [
            ('id-a', 'a@example.invalid', 'never-return-hash'),
            ('id-b', 'b@example.invalid', 'never-return-hash'),
            ('unlinked', 'unlinked@example.invalid', 'never-return-hash'),
        ])

    def register(self):
        self.db.execute('''CREATE TABLE personnel_register(employee_id INTEGER PRIMARY KEY,
            linked_username TEXT UNIQUE,full_name TEXT,address TEXT,birth_date TEXT,
            social_security_number_encrypted TEXT)''')

    def test_missing_register_is_backward_compatible(self):
        self.assertEqual(load_person_names(self.db.cursor()), {})
        accounts = load_person_accounts(self.db.cursor())
        self.assertEqual(accounts['id-a']['display_name'], 'a')

    def test_only_explicit_links_are_used_even_when_names_look_like_ids(self):
        self.register()
        self.db.executemany('INSERT INTO personnel_register VALUES(?,?,?,?,?,?)', [
            (1, 'id-a', 'Same Name', 'private-address', '2000-01-01', 'private-envelope'),
            (2, 'id-b', 'Same Name', 'private-address', '2000-01-01', 'private-envelope'),
            (3, None, 'unlinked', 'private-address', '2000-01-01', 'private-envelope'),
        ])
        self.assertEqual(load_person_names(self.db.cursor()), {'id-a': 'Same Name', 'id-b': 'Same Name'})
        accounts = load_person_accounts(self.db.cursor())
        self.assertEqual(accounts['unlinked']['display_name'], 'unlinked')
        self.assertIsNone(accounts['unlinked']['full_name'])

    def test_sensitive_columns_and_credentials_are_not_selected(self):
        self.register()
        self.db.execute('INSERT INTO personnel_register VALUES(1,?,?,?,?,?)',
                        ('id-a', 'Visible Name', 'private-address', '2000-01-01', 'private-envelope'))
        sql = []
        self.db.set_trace_callback(sql.append)
        self.addCleanup(lambda: self.db.set_trace_callback(None))
        result = load_person_accounts(self.db.cursor(), ['id-a'])
        self.assertEqual(set(result['id-a']), {'username', 'email', 'full_name', 'name', 'display_name'})
        rendered = json.dumps(result)
        for secret in ('private-address', '2000-01-01', 'private-envelope', 'never-return-hash'):
            self.assertNotIn(secret, rendered)
        for statement in sql:
            self.assertNotIn('*', statement)
            self.assertFalse(re.search(r'password|social_security|birth_date|address', statement, re.I))

    def test_tuple_dbapi_rows_are_supported(self):
        self.register()
        self.db.execute('INSERT INTO personnel_register(employee_id,linked_username,full_name) VALUES(1,?,?)',
                        ('id-a', 'Visible Name'))
        self.db.row_factory = None
        self.assertEqual(load_person_names(self.db.cursor()), {'id-a': 'Visible Name'})
        self.assertEqual(load_display_names(self.db.cursor(), ['id-a']), {'id-a': 'Visible Name'})

    def test_sql_parameterization_and_deduplication(self):
        self.register()
        supplied = ["id-a') OR 1=1 --", 'id-a', 'id-a', None, 2]
        self.db.execute('INSERT INTO personnel_register(employee_id,linked_username,full_name) VALUES(1,?,?)',
                        ('id-b', 'Hidden Outside Selection'))
        self.assertEqual(load_person_names(self.db.cursor(), supplied), {})
        self.assertEqual(set(load_person_accounts(self.db.cursor(), supplied)), {'id-a'})

    def test_empty_selection_does_not_query(self):
        cursor = Mock()
        self.assertEqual(load_person_names(cursor, []), {})
        self.assertEqual(load_person_accounts(cursor, []), {})
        self.assertEqual(load_display_names(cursor, []), {})
        cursor.execute.assert_not_called()

    def test_batches_larger_than_sqlite_legacy_parameter_limit(self):
        self.register()
        accounts = [(f'synthetic-{i}', f'prefix{i}@example.invalid', 'hash') for i in range(1100)]
        self.db.executemany('INSERT INTO users VALUES(?,?,?)', accounts)
        self.db.executemany('INSERT INTO personnel_register(employee_id,linked_username,full_name) VALUES(?,?,?)',
                            [(i + 1, row[0], f'Person {i}') for i, row in enumerate(accounts)])
        names = load_person_names(self.db.cursor(), [row[0] for row in accounts])
        self.assertEqual(len(names), 1100)
        self.assertEqual(names['synthetic-1099'], 'Person 1099')
        self.assertEqual(len(load_person_accounts(self.db.cursor(), [row[0] for row in accounts])), 1100)

    def test_read_queries_do_not_rewrite_stored_names_or_links(self):
        self.register()
        self.db.execute('INSERT INTO personnel_register(employee_id,linked_username,full_name) VALUES(1,?,?)',
                        ('id-a', '  Original Name  '))
        before = list(self.db.execute('SELECT * FROM personnel_register'))
        load_person_accounts(self.db.cursor())
        self.assertEqual(list(self.db.execute('SELECT * FROM personnel_register')), before)

    def test_mysql_missing_table_race_is_safe_but_other_failures_propagate(self):
        cursor = Mock()
        cursor.fetchone.return_value = {'present': True}
        cursor.execute.side_effect = [None, ProgrammingError(1146, 'Synthetic absent optional table')]
        self.assertEqual(load_person_names(cursor, ['id-a']), {})
        cursor.execute.side_effect = [None, ProgrammingError(1054, 'Synthetic bad column')]
        with self.assertRaises(ProgrammingError): load_person_names(cursor, ['id-a'])
        cursor.execute.side_effect = OperationalError(2006, 'Synthetic connection lost')
        with self.assertRaises(OperationalError): load_person_names(cursor, ['id-a'])


class NameCursor:
    """Translate syntax only; preserve actual SQLite rows/filters and DB-API metadata."""
    def __init__(self, db):
        self.connection = db
        self.cursor = db.cursor()
    def __enter__(self): return self
    def __exit__(self, *_): self.cursor.close()
    def execute(self, sql, params=()):
        match = re.fullmatch(r"SHOW TABLES LIKE '([A-Za-z_]+)'", sql)
        if match:
            return self.cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name=?", (match[1],))
        return self.cursor.execute(sql.replace('%s', '?').replace(' FOR UPDATE', ''), params)
    def fetchone(self):
        row = self.cursor.fetchone()
        return dict(row) if row is not None else None
    def fetchall(self): return [dict(row) for row in self.cursor.fetchall()]
    @property
    def description(self): return self.cursor.description
    @property
    def lastrowid(self): return self.cursor.lastrowid


class NameConnection:
    def __init__(self, path):
        self.db = sqlite3.connect(path)
        self.db.row_factory = sqlite3.Row
        self.db.create_function('UNIX_TIMESTAMP', 1, lambda value:
            datetime.fromisoformat(value).replace(tzinfo=timezone.utc).timestamp() if value else None)
    def cursor(self): return NameCursor(self.db)
    def commit(self): self.db.commit()
    def rollback(self): self.db.rollback()
    def close(self): self.db.close()


class PersonNameProjectionHTTPTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.path = Path(self.temp.name) / 'name-projection.sqlite3'
        with closing(sqlite3.connect(self.path)) as db, db:
            db.executescript('''CREATE TABLE users(username TEXT PRIMARY KEY,email TEXT,
                password_hash TEXT DEFAULT 'private-hash',is_admin INTEGER DEFAULT 0,
                role TEXT DEFAULT 'Employe',role_tags TEXT,employee_type TEXT DEFAULT 'salarie',
                manager_username TEXT,contract_start_date TEXT,contract_end_date TEXT,is_cdi INTEGER DEFAULT 0,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP,is_active INTEGER DEFAULT 1,
                session_version INTEGER DEFAULT 0,login_deleted_at TEXT);
                CREATE TABLE personnel_register(employee_id INTEGER PRIMARY KEY AUTOINCREMENT,
                full_name TEXT,nationality TEXT DEFAULT '',sex TEXT DEFAULT '',address TEXT DEFAULT '',
                birth_date TEXT,job_title TEXT DEFAULT '',qualification TEXT DEFAULT '',personnel_type TEXT DEFAULT '',
                linked_username TEXT UNIQUE,social_security_number_encrypted TEXT,is_archived INTEGER DEFAULT 0,
                version INTEGER DEFAULT 1,created_at TEXT DEFAULT CURRENT_TIMESTAMP,updated_at TEXT DEFAULT CURRENT_TIMESTAMP);
                CREATE TABLE personnel_register_periods(period_id INTEGER PRIMARY KEY,employee_id INTEGER,
                entry_date TEXT,exit_date TEXT,contract_type TEXT DEFAULT '',other_contract TEXT DEFAULT '',
                created_at TEXT DEFAULT CURRENT_TIMESTAMP,updated_at TEXT DEFAULT CURRENT_TIMESTAMP);
                CREATE TABLE personnel_register_events(event_id INTEGER PRIMARY KEY,employee_id INTEGER,
                actor_username TEXT,action TEXT,before_snapshot TEXT,after_snapshot TEXT,created_at TEXT DEFAULT CURRENT_TIMESTAMP);
                CREATE TABLE organization_groups(id INTEGER PRIMARY KEY,name TEXT,is_active INTEGER DEFAULT 1);
                CREATE TABLE user_group_memberships(username TEXT,group_id INTEGER,relation TEXT);
                CREATE TABLE transport_receipts(employee_username TEXT,receipt_month INTEGER,receipt_year INTEGER,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP,archived INTEGER DEFAULT 0);
                CREATE TABLE leave_requests(id INTEGER PRIMARY KEY,employee_username TEXT,start_date TEXT,
                end_date TEXT,start_slot TEXT,end_slot TEXT,status TEXT);
                CREATE TABLE Presence(user_id TEXT,form_content TEXT);''')
            for username, role, manager in [('admin-id', 'Admin', None), ('manager-id', 'Responsable', None),
                    ('person-a', 'Employe', 'manager-id'), ('person-b', 'Employe', 'admin-id'),
                    ('outsider-id', 'Employe', None), ('unlinked-id', 'Employe', None)]:
                db.execute('INSERT INTO users(username,email,role,role_tags,manager_username) VALUES(?,?,?,?,?)',
                           (username, username + '@example.invalid', role, json.dumps([role]), manager))
            for username, name in [('admin-id', 'Synthetic Admin'), ('manager-id', 'Synthetic Manager'),
                                   ('person-a', 'Namesake Person'), ('person-b', 'Namesake Person'),
                                   ('outsider-id', 'Namesake Person')]:
                db.execute('''INSERT INTO personnel_register(full_name,linked_username,address,birth_date,
                    social_security_number_encrypted) VALUES(?,?,?,?,?)''',
                    (name, username, 'private-address-marker', '2000-01-01', 'private-envelope-marker'))
            db.execute("INSERT INTO organization_groups VALUES(1,'Synthetic Department',1)")
            for username in ('person-a', 'person-b'):
                db.execute('INSERT INTO user_group_memberships VALUES(?,1,?)', (username, 'member'))
            db.execute('INSERT INTO user_group_memberships VALUES(?,1,?)', ('manager-id', 'manager'))
            snapshot = {'employee': {'employee_id': 3, 'full_name': 'Historical Name',
                        'linked_username': 'person-a', 'has_social_security_number': True}, 'periods': []}
            db.execute('INSERT INTO personnel_register_events VALUES(1,3,?,?,NULL,?,CURRENT_TIMESTAMP)',
                        ('manager-id', 'created', json.dumps(snapshot)))
        self.actor = {'username': 'person-a', 'role_tags': ['Employe'], 'session_version': 0}
        def authenticated():
            if self.actor is None: raise HTTPException(401, 'Authentication required')
            return dict(self.actor)
        self.connect = lambda: NameConnection(self.path)
        self.attempts = Mock()
        self.attempts.profile_receipts.return_value = ([], True)
        app = FastAPI()
        register_team_directory(app, get_db_connection=self.connect, get_current_user=authenticated)
        register_profile_routes(app, get_db_connection=self.connect, get_current_user=authenticated,
                                load_user=test_backend_helpers.main.load_user, attempts=self.attempts)
        register_personnel_routes(app, get_db_connection=self.connect, get_current_user=authenticated,
                                  load_user=test_backend_helpers.main.load_user)
        self.client = TestClient(app)
        self.addCleanup(self.client.close)

    def test_public_user_keeps_identity_and_hides_private_source_columns(self):
        connection = self.connect()
        try:
            with connection.cursor() as cursor:
                user = test_backend_helpers.main.load_user(cursor, 'person-a')
                public = test_backend_helpers.main.public_user(user)
        finally: connection.close()
        self.assertEqual(public['display_name'], 'Namesake Person')
        self.assertEqual(public['username'], 'person-a')
        self.assertEqual(public['id'], 'person-a')
        self.assertEqual(public['manager_id'], 'manager-id')
        self.assertEqual(public['manager_display_name'], 'Synthetic Manager')
        for field in ('password_hash', 'social_security_number_encrypted', 'address', 'birth_date'):
            self.assertNotIn(field, public)

    def test_account_directory_keeps_namesakes_as_separate_canonical_accounts(self):
        connection = self.connect()
        try:
            with connection.cursor() as cursor:
                users = test_backend_helpers.main.all_users(cursor)
        finally: connection.close()
        self.assertEqual(len(users), 6)
        by_id = {user['username']: test_backend_helpers.main.public_user(user) for user in users}
        self.assertEqual(by_id['person-a']['display_name'], by_id['person-b']['display_name'])
        self.assertNotEqual(by_id['person-a']['manager_id'], by_id['person-b']['manager_id'])
        self.assertEqual(by_id['unlinked-id']['display_name'], 'unlinked-id')
        self.assertNotIn('private-hash', json.dumps(by_id))

    def test_organization_uses_names_without_email_or_sensitive_register_fields(self):
        response = self.client.get('/organization-directory')
        self.assertEqual(response.status_code, 200, response.text)
        users = {row['username']: row for row in response.json()['users']}
        self.assertEqual(users['person-a']['display_name'], 'Namesake Person')
        self.assertEqual(users['unlinked-id']['display_name'], 'unlinked-id')
        for field in ('email', 'address', 'birth_date', 'social_security_number_encrypted', 'password_hash'):
            self.assertTrue(all(field not in row for row in users.values()))
        self.assertNotIn('private-envelope-marker', response.text)

    def test_namesakes_remain_two_members_and_query_cannot_expand_team_scope(self):
        response = self.client.get('/my-team?username=outsider-id&group_id=999')
        self.assertEqual(response.status_code, 200, response.text)
        members = response.json()['members']
        self.assertEqual(len([row for row in members if row['name'] == 'Namesake Person']), 2)
        self.assertEqual(len(members), 3)
        self.assertNotIn('username', members[0])
        self.assertNotIn('email', members[0])

    def test_duplicate_membership_rows_deduplicate_identity_not_display_name(self):
        with closing(sqlite3.connect(self.path)) as db, db:
            db.execute('INSERT INTO user_group_memberships VALUES(?,1,?)', ('person-a', 'member'))
        response = self.client.get('/my-team')
        self.assertEqual(response.status_code, 200, response.text)
        self.assertEqual(len(response.json()['members']), 3)
        self.assertEqual(sum(row['name'] == 'Namesake Person' for row in response.json()['members']), 2)

    def test_same_display_name_does_not_authorize_other_profile(self):
        self.assertEqual(self.client.get('/person-profiles/person-b?year=2026').status_code, 403)
        self.assertEqual(self.client.get('/person-profiles/Namesake%20Person?year=2026').status_code, 403)
        self.attempts.profile_receipts.assert_not_called()

    def test_profile_access_uses_ids_for_self_manager_and_admin(self):
        for actor in ('person-a', 'manager-id', 'admin-id'):
            with self.subTest(actor=actor):
                self.actor['username'] = actor
                response = self.client.get('/person-profiles/person-a?year=2026')
                self.assertEqual(response.status_code, 200, response.text)
                data = response.json()
                self.assertEqual(data['username'], 'person-a')
                self.assertEqual(data['name'], 'Namesake Person')
                self.assertEqual(data['manager'], 'manager-id')
                self.assertEqual(data['manager_name'], 'Synthetic Manager')
                self.assertNotIn('private-envelope-marker', response.text)
                self.assertNotIn('private-address-marker', response.text)
        self.assertEqual(self.attempts.profile_receipts.call_args.args, ('person-a', 2026))

    def test_plain_employee_cannot_search_register_names_or_sensitive_details(self):
        for path in ('/personnel-register/accounts?q=Namesake', '/personnel-register/3',
                     '/personnel-register/3/history'):
            with self.subTest(path=path):
                response = self.client.get(path)
                self.assertEqual(response.status_code, 403)
                self.assertNotIn('private-envelope-marker', response.text)

    def test_admin_account_search_finds_names_but_does_not_dedupe_namesakes(self):
        self.actor['username'] = 'admin-id'
        response = self.client.get('/personnel-register/accounts?q=Namesake')
        self.assertEqual(response.status_code, 200, response.text)
        accounts = response.json()['accounts']
        self.assertEqual({row['username'] for row in accounts}, {'person-a', 'person-b', 'outsider-id'})
        self.assertEqual({row['display_name'] for row in accounts}, {'Namesake Person'})
        self.assertTrue(all('social_security_number_encrypted' not in row for row in accounts))
        self.assertTrue(all('address' not in row for row in accounts))

    def test_history_projects_actor_but_keeps_recorded_ids_and_names(self):
        self.actor['username'] = 'admin-id'
        with closing(sqlite3.connect(self.path)) as db:
            before = db.execute('SELECT after_snapshot FROM personnel_register_events').fetchone()[0]
        response = self.client.get('/personnel-register/3/history')
        self.assertEqual(response.status_code, 200, response.text)
        event = response.json()['events'][0]
        self.assertEqual(event['actor_username'], 'manager-id')
        self.assertEqual(event['actor_display_name'], 'Synthetic Manager')
        self.assertEqual(event['after']['employee']['linked_username'], 'person-a')
        self.assertEqual(event['after']['employee']['full_name'], 'Historical Name')
        with closing(sqlite3.connect(self.path)) as db:
            self.assertEqual(db.execute('SELECT after_snapshot FROM personnel_register_events').fetchone()[0], before)

    def test_fresh_requests_reflect_updated_linked_name_without_changing_ids(self):
        self.assertEqual(self.client.get('/person-profiles/person-a?year=2026').json()['name'], 'Namesake Person')
        with closing(sqlite3.connect(self.path)) as db, db:
            db.execute('UPDATE personnel_register SET full_name=? WHERE linked_username=?',
                       ('Updated Synthetic Name', 'person-a'))
        response = self.client.get('/person-profiles/person-a?year=2026')
        self.assertEqual(response.json()['name'], 'Updated Synthetic Name')
        self.assertEqual(response.json()['username'], 'person-a')
        self.assertEqual(self.client.get('/person-profiles/person-b?year=2026').status_code, 403)


if __name__ == '__main__':
    unittest.main()
