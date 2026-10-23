"""Account identity storage and name projections on synthetic isolated fixtures."""
import sqlite3
import os
import unittest
from unittest.mock import Mock
from fastapi import HTTPException
from pydantic import ValidationError
from backend.account_names import AccountNames, ensure_account_names_schema
from backend.person_names import load_person_accounts, load_person_names
import test_super_admin_roles as fixture
import test_personnel_sensitive_mysql as mysql_fixture


class AccountNameInputTests(unittest.TestCase):
    def test_required_names_preserve_unicode_and_compound_surnames(self):
        names = AccountNames(first_name='  Élodie   Anne ', last_name="  de La Tour-张  ")
        self.assertEqual(names.full_name, 'Élodie Anne de La Tour-张')
        for values in ({'first_name': '', 'last_name': 'Nom'}, {'first_name': 'Nom'},
                       {'first_name': 'Prénom', 'last_name': '  '}, {'first_name': 12, 'last_name': 'Nom'},
                       {'first_name': 'A' * 100, 'last_name': 'B' * 100},
                       {'first_name': 'A\x00B', 'last_name': 'Nom'}):
            with self.subTest(values=values), self.assertRaises(ValidationError):
                AccountNames(**values)

    def test_migration_is_additive_and_restores_session_sql_mode_on_failure(self):
        cursor = Mock()
        ensure_account_names_schema(cursor, {'first_name', 'last_name'})
        cursor.execute.assert_not_called()
        cursor.fetchone.return_value = {'sql_mode': 'STRICT_TRANS_TABLES,NO_ZERO_DATE,NO_ZERO_IN_DATE'}
        ensure_account_names_schema(cursor, set())
        commands = [call.args[0] for call in cursor.execute.call_args_list]
        self.assertEqual(sum(sql.startswith('ALTER TABLE users ADD COLUMN') for sql in commands), 1)
        self.assertFalse(any(sql.startswith(('UPDATE ', 'DELETE ', 'DROP ')) for sql in commands))
        cursor.reset_mock()
        def fail(sql, *_):
            if sql.startswith('ALTER TABLE'): raise RuntimeError('Synthetic ALTER failure')
        cursor.execute.side_effect = fail
        with self.assertRaises(RuntimeError): ensure_account_names_schema(cursor, set())
        self.assertEqual(cursor.execute.call_args.args, ('SET SESSION sql_mode=%s',
                         ('STRICT_TRANS_TABLES,NO_ZERO_DATE,NO_ZERO_IN_DATE',)))


class AccountCreationNameTests(unittest.TestCase):
    setUp = fixture.MainRoleMutationTests.setUp

    @staticmethod
    def payload(**changes):
        return {'first_name': '  Élodie  Anne ', 'last_name': 'Durand',
                'new_mail': 'synthetic-new@example.invalid', 'new_password': 'synthetic-unit-only',
                'role_tags': ['Employe'], 'manager_id': 'admin', **changes}

    def test_create_persists_separate_names_with_opaque_id_and_keeps_old_accounts(self):
        before = set(self.fixture.users)
        result = fixture.main.create_user(self.payload(), self.fixture.actor())
        self.assertEqual(result['status'], 'success')
        key, = set(self.fixture.users) - before
        row = self.fixture.users[key]
        self.assertRegex(key, r'^u_[0-9a-f]{24}$')
        self.assertEqual(row['first_name'], 'Élodie Anne')
        self.assertEqual(row['last_name'], 'Durand')
        self.assertEqual(row['email'], 'synthetic-new@example.invalid')
        self.assertEqual(row['password_hash'], 'synthetic-hash')
        self.assertEqual(set(self.fixture.users) & before, before)

    def test_missing_names_rejected_before_any_database_write(self):
        before = set(self.fixture.users)
        for key in ('first_name', 'last_name'):
            with self.subTest(key=key), self.assertRaises(HTTPException) as error:
                fixture.main.create_user(self.payload(**{key: ' '}), self.fixture.actor())
            self.assertEqual(error.exception.status_code, 400)
        self.assertEqual(set(self.fixture.users), before)
        self.assertFalse(self.fixture.statements)

    def test_role_audit_failure_rolls_back_identity_with_account(self):
        self.fixture.fail_audit = True
        with self.assertRaises(RuntimeError): fixture.main.create_user(self.payload(), self.fixture.actor())
        self.assertEqual(set(self.fixture.users), {'root', 'admin', 'worker'})
        self.assertEqual(self.fixture.rollbacks, 1)


class AccountNameProjectionTests(unittest.TestCase):
    def setUp(self):
        self.db = sqlite3.connect(':memory:')
        self.db.row_factory = sqlite3.Row
        self.addCleanup(self.db.close)
        self.db.execute("CREATE TABLE users(username TEXT PRIMARY KEY,email TEXT,first_name TEXT,last_name TEXT,password_hash TEXT)")
        self.db.execute('INSERT INTO users VALUES(?,?,?,?,?)',
                        ('opaque-id', 'prefix@example.invalid', 'Élodie Anne', 'Durand', 'private-hash'))

    def test_account_name_works_without_register_and_contains_no_credentials(self):
        values = load_person_accounts(self.db.cursor())
        self.assertEqual(values['opaque-id']['display_name'], 'Élodie Anne Durand')
        self.assertNotIn('password_hash', values['opaque-id'])
        self.assertNotIn('opaque-id', values['opaque-id']['display_name'])

    def test_registered_structured_name_wins_but_legacy_label_cannot_override_account_names(self):
        self.db.execute('CREATE TABLE personnel_register(linked_username TEXT,full_name TEXT,first_name TEXT,last_name TEXT)')
        self.db.execute('INSERT INTO personnel_register VALUES(?,?,?,?)',
                        ('opaque-id', 'Unstructured legacy label', '', ''))
        self.assertEqual(load_person_names(self.db.cursor())['opaque-id'], 'Élodie Anne Durand')
        self.db.execute('UPDATE personnel_register SET first_name=?,last_name=?', ('Élodie', 'Martin'))
        self.assertEqual(load_person_names(self.db.cursor())['opaque-id'], 'Élodie Martin')
        self.db.execute('UPDATE personnel_register SET linked_username=NULL')
        self.assertEqual(load_person_names(self.db.cursor())['opaque-id'], 'Élodie Anne Durand')

    def test_namesakes_keep_distinct_account_keys(self):
        self.db.execute('INSERT INTO users VALUES(?,?,?,?,?)',
                        ('another-id', 'other@example.invalid', 'Élodie Anne', 'Durand', 'private-hash'))
        values = load_person_accounts(self.db.cursor())
        self.assertEqual(set(values), {'opaque-id', 'another-id'})
        self.assertEqual(values['opaque-id']['display_name'], values['another-id']['display_name'])


@unittest.skipUnless(os.getenv('PRESENCE_MYSQL_TEMP_TEST') == '1', 'opt-in MySQL temporary-table integration')
class AccountNameMySQLTests(unittest.TestCase):
    setUp = mysql_fixture.PersonnelSensitiveMySQLTests.setUp

    def test_additive_migration_preserves_old_users_and_new_names_round_trip(self):
        # This connection owns a TEMPORARY users table shadowing the real one.
        with self.db.cursor() as cursor:
            cursor.execute('SELECT * FROM users ORDER BY username')
            before = cursor.fetchall()
            cursor.execute('SHOW COLUMNS FROM users')
            ensure_account_names_schema(cursor, {row['Field'] for row in cursor.fetchall()})
            cursor.execute('SHOW COLUMNS FROM users')
            ensure_account_names_schema(cursor, {row['Field'] for row in cursor.fetchall()})
            cursor.execute('SELECT * FROM users ORDER BY username')
            after = cursor.fetchall()
            self.assertEqual([{key: row[key] for key in before[0]} for row in after], before)
            self.assertTrue(all(row['first_name'] == row['last_name'] == '' for row in after))
            cursor.execute('INSERT INTO users(username,email,role_tags,first_name,last_name) VALUES(%s,%s,%s,%s,%s)',
                           ('opaque-synthetic', 'synthetic@example.invalid', '["Employe"]', 'Élodie Anne', 'Durand-张'))
            values = load_person_accounts(cursor, ['opaque-synthetic'])
            self.assertEqual(values['opaque-synthetic']['display_name'], 'Élodie Anne Durand-张')
            self.assertNotIn('role_tags', values['opaque-synthetic'])
