"""Mailbox identity and release rules on owned MySQL TEMPORARY tables only."""
import json
import os
import unittest
from unittest.mock import patch

import pymysql
from fastapi import HTTPException
from backend.account_email import ensure_account_email_schema
from backend.login_guard import ensure_login_guard_schema
from backend.signature_flow import resolve_manager
from test_support.personnel_fixtures import registration_payload
import test_responsable_workflow_mysql as workflow


@unittest.skipUnless(os.getenv('PRESENCE_MYSQL_TEMP_TEST') == '1', 'owned MySQL temporary tables')
class AccountEmailTests(unittest.TestCase):
    request = workflow.ResponsableWorkflowMySQLTests.request

    def setUp(self):
        workflow.ResponsableWorkflowMySQLTests.setUp(self)
        with self.db.cursor() as cursor:
            cursor.execute('ALTER TABLE users ADD UNIQUE KEY email (email)')
            ensure_account_email_schema(cursor)
            ensure_account_email_schema(cursor)
            cursor.execute('SHOW COLUMNS FROM users')
            ensure_login_guard_schema(cursor, {r['Field'] for r in cursor.fetchall()})
        self.db.commit()

    def create_account(self, email='mailbox@example.com'):
        self.request('POST', '/create-user', json={
            'first_name': 'Synthetic', 'last_name': 'Mailbox', 'new_mail': email,
            'new_password': 'Synthetic8', 'role_tags': ['Employe']})
        with self.db.cursor() as c:
            c.execute('SELECT username FROM users WHERE email=%s AND login_deleted_at IS NULL', (email,))
            return c.fetchone()['username']

    def register(self, username, email='mailbox@example.com'):
        return self.request('POST', '/personnel-register', status=201,
            json=registration_payload(linked_username=username, email=email)).json()['employee']

    def update_email(self, employee, email, status=200):
        body = registration_payload(linked_username=employee['linked_username'], email=email)
        body.pop('initial_period')
        return self.request('PUT', f"/personnel-register/{employee['employee_id']}", status=status,
                            json={**body, 'expected_version': employee['version']})

    def test_disabled_and_archived_accounts_keep_email_case_insensitively(self):
        name = self.create_account()
        employee = self.register(name)
        with self.db.cursor() as c:
            c.execute('UPDATE users SET is_active=0 WHERE username=%s', (name,))
        self.db.commit()
        self.request('PATCH', f"/personnel-register/{employee['employee_id']}/archive",
                     json={'is_archived': True, 'expected_version': employee['version']})
        self.request('POST', '/create-user', status=400, json={
            'first_name': 'Other', 'last_name': 'Fixture', 'new_mail': ' MAILBOX@example.com ',
            'new_password': 'Synthetic8', 'role_tags': ['Employe']})
        result = self.request('GET', '/personnel-register/email-availability?email=MAILBOX@example.com').json()
        self.assertFalse(result['available'])

    def test_database_rejects_direct_duplicate_even_without_application_check(self):
        self.create_account()
        with self.db.cursor() as c:
            with self.assertRaises(pymysql.err.IntegrityError):
                c.execute('INSERT INTO users(username,email) VALUES(%s,%s)', ('direct-fixture', ' MAILBOX@example.com '))
            c.execute('SHOW INDEX FROM users')
            names = {r['Key_name'] for r in c.fetchall()}
        self.assertIn('users_login_email_unique', names)
        self.assertNotIn('email', names)

    def test_legacy_zero_date_default_migration_restores_session_sql_mode(self):
        with self.db.cursor() as c:
            c.execute('SELECT @@SESSION.sql_mode AS value')
            previous = c.fetchone()['value']
            relaxed = ','.join(p for p in previous.split(',') if p not in {'NO_ZERO_DATE', 'NO_ZERO_IN_DATE'})
            c.execute('SET SESSION sql_mode=%s', (relaxed,))
            c.execute('ALTER TABLE users DROP INDEX users_login_email_unique, DROP COLUMN login_email_key')
            c.execute("ALTER TABLE users MODIFY reset_token_expires DATETIME DEFAULT '0000-00-00 00:00:00'")
            c.execute('SET SESSION sql_mode=%s', (previous,))
            ensure_account_email_schema(c)
            c.execute('SELECT @@SESSION.sql_mode AS value')
            self.assertEqual(c.fetchone()['value'], previous)

    def test_admin_cannot_change_hidden_super_admin_login_through_register(self):
        employee = self.register('fixture-employee', 'fixture-contact@example.com')
        with self.db.cursor() as c:
            c.execute("UPDATE users SET role='SuperAdmin',role_tags=%s,is_admin=1 WHERE username='fixture-employee'",
                      ('["SuperAdmin","Admin"]',))
        self.db.commit()
        self.update_email(employee, 'blocked@example.com', status=403)
        with self.db.cursor() as c:
            c.execute("SELECT session_version FROM users WHERE username='fixture-employee'")
            self.assertEqual(c.fetchone()['session_version'], 0)

    def test_deleted_login_releases_mail_without_rewriting_history_or_ambiguating_login(self):
        old = self.create_account()
        employee = self.register(old)
        with self.db.cursor() as c:
            c.execute('UPDATE users SET is_active=0,login_deleted_at=CURRENT_TIMESTAMP WHERE username=%s', (old,))
        self.db.commit()
        new = self.create_account()
        self.assertNotEqual(new, old)
        with self.db.cursor() as c:
            c.execute('SELECT email FROM users WHERE username=%s', (old,))
            self.assertEqual(c.fetchone()['email'], 'mailbox@example.com')
        saved = self.request('GET', f"/personnel-register/{employee['employee_id']}").json()['employee']
        self.assertEqual(saved['linked_username'], old)
        self.assertEqual(saved['email'], 'mailbox@example.com')
        result = workflow.main.login({'username': 'MAILBOX@example.com', 'password': 'Synthetic8'})
        self.assertEqual(result['username'], new)
        self.request('PUT', f'/users/{new}/organization', json={'role_tags': ['Employe'], 'manager_id': 'fixture-manager'})
        manager = resolve_manager('mailbox@example.com', workflow.main.load_user(self.proxy.cursor(), 'fixture-admin'),
                                  lambda: self.proxy, workflow.main.load_user, workflow.main.require_target_access,
                                  include_identity=True, employee_username=new)
        self.assertEqual(manager['employee_username'], new)
        with patch.object(workflow.main, 'send_password_reset_email') as send:
            workflow.main.forgot_password({'email': 'mailbox@example.com'})
            send.assert_called_once()
        with self.db.cursor() as c:
            c.execute('SELECT username,reset_token FROM users WHERE username IN (%s,%s)', (old, new))
            tokens = {row['username']: row['reset_token'] for row in c.fetchall()}
        self.assertIsNone(tokens[old])
        self.assertTrue(tokens[new])

    def test_register_email_changes_login_and_revokes_old_sessions(self):
        name = self.create_account()
        employee = self.register(name)
        updated = self.update_email(employee, 'changed@example.com').json()['employee']
        self.assertEqual(updated['email'], 'changed@example.com')
        old = workflow.main.create_session_token(name, 0)
        response = self.client.get('/account', headers={'Authorization': 'Bearer ' + old})
        self.assertEqual(response.status_code, 401)
        result = workflow.main.login({'username': 'changed@example.com', 'password': 'Synthetic8'})
        self.assertEqual(result['username'], name)
        with self.assertRaises(HTTPException) as denied:
            workflow.main.login({'username': 'mailbox@example.com', 'password': 'Synthetic8'})
        self.assertEqual(denied.exception.status_code, 401)

    def test_duplicate_or_empty_register_email_rolls_back_both_records(self):
        name = self.create_account()
        employee = self.register(name)
        self.create_account('other@example.com')
        self.update_email(employee, 'OTHER@example.com', status=400)
        self.update_email(employee, '', status=400)
        saved = self.request('GET', f"/personnel-register/{employee['employee_id']}").json()['employee']
        self.assertEqual(saved['version'], employee['version'])
        self.assertEqual(saved['email'], 'mailbox@example.com')

    def test_audit_failure_rolls_back_linked_login_and_mail(self):
        name = self.create_account()
        employee = self.register(name)
        self.proxy.fail_audit = True
        self.update_email(employee, 'changed@example.com', status=500)
        with self.db.cursor() as c:
            c.execute('SELECT email,session_version FROM users WHERE username=%s', (name,))
            saved = c.fetchone()
        self.assertEqual(saved['email'], 'mailbox@example.com')
        self.assertEqual(saved['session_version'], 0)

    def test_unlinked_contact_email_does_not_change_any_login(self):
        self.create_account()
        employee = self.request('POST', '/personnel-register', status=201,
            json=registration_payload(email='contact@example.com')).json()['employee']
        self.update_email(employee, 'mailbox@example.com')
        self.request('GET', '/personnel-register/email-availability?email=mailbox@example.com', actor='fixture-employee', status=403)

    def test_same_own_email_available_but_profile_cannot_take_another_login_mail(self):
        name = self.create_account()
        employee = self.register(name)
        self.assertTrue(self.request('GET', f"/personnel-register/email-availability?email=MAILBOX@example.com&employee_id={employee['employee_id']}").json()['available'])
        self.create_account('other@example.com')
        with patch.object(workflow.main, 'verify_password', return_value=True):
            with self.assertRaises(HTTPException) as error:
                workflow.main.update_profile(name, workflow.main.ProfileUpdateRequest(current_password='synthetic', new_email='other@example.com'),
                                             workflow.main.load_user(self.proxy.cursor(), name))
        self.assertEqual(error.exception.status_code, 400)

    def test_profile_email_updates_linked_register_and_returns_new_session(self):
        name = self.create_account()
        employee = self.register(name)
        with patch.object(workflow.main, 'verify_password', return_value=True), patch.object(workflow.main, 'migrate_username_in_saved_fiches'):
            result = workflow.main.update_profile(name, workflow.main.ProfileUpdateRequest(current_password='synthetic', new_email='profile@example.com'),
                                                 workflow.main.load_user(self.proxy.cursor(), name))
        token = workflow.main.read_session_token(result['auth_token'], with_version=True)
        self.assertEqual(token['session_version'], 1)
        saved = self.request('GET', f"/personnel-register/{employee['employee_id']}").json()['employee']
        self.assertEqual(saved['email'], 'profile@example.com')
        self.assertGreater(saved['version'], employee['version'])
