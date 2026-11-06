"""Atomic employee/account registration on disposable MySQL tables; SMTP mocked."""
import json
import os
import unittest
from pydantic import ValidationError
from backend.personnel_rules import EmployeeCreate
from test_support.personnel_fixtures import registration_payload
import test_responsable_workflow_mysql as workflow


class RegistrationAccountRulesTests(unittest.TestCase):
    def test_email_optional_without_account_but_valid_when_provided(self):
        self.assertEqual(EmployeeCreate(**registration_payload()).email, '')
        self.assertEqual(EmployeeCreate(**registration_payload(email='test@example.com')).email, 'test@example.com')
        with self.assertRaises(ValidationError):
            EmployeeCreate(**registration_payload(email='invalid'))

    def test_new_account_requires_email_and_excludes_existing_link(self):
        for changes in ({'account': {}}, {'account': {}, 'email': 'test@example.com', 'linked_username': 'existing'}):
            with self.subTest(changes=changes), self.assertRaises(ValidationError):
                EmployeeCreate(**registration_payload(**changes))

    def test_client_cannot_supply_password_or_account_identity(self):
        for field in ('new_password', 'new_username', 'generate_password'):
            with self.subTest(field=field), self.assertRaises(ValidationError):
                EmployeeCreate(**registration_payload(email='test@example.com', account={field: 'forged'}))


@unittest.skipUnless(os.getenv('PRESENCE_MYSQL_TEMP_TEST') == '1', 'opt-in loopback temporary tables')
class RegistrationAccountMySQLTests(unittest.TestCase):
    setUp = workflow.ResponsableWorkflowMySQLTests.setUp
    request = workflow.ResponsableWorkflowMySQLTests.request

    def payload(self, **changes):
        return registration_payload(email='new-register@example.com',
            account={'role_tags': ['Employe'], 'group_ids': [1], 'manager_id': 'fixture-manager'},
            initial_period={'entry_date': '2026-01-01', 'contract_type': 'CDI'}, **changes)

    def count(self, table):
        self.assertIn(table, {'users', 'personnel_register', 'personnel_contracts', 'account_role_events', 'user_group_memberships'})
        with self.db.cursor() as cursor:
            cursor.execute(f'SELECT COUNT(*) AS n FROM {table}')
            return cursor.fetchone()['n']

    def test_atomic_registration_links_account_contract_and_hashed_password(self):
        def mail(email, name, password, generated):
            self.assertFalse(self.db.server_status & 1, 'SMTP must follow commit')
            self.assertEqual(len(password), 8)
            self.assertTrue(generated)
            self.assertEqual(email, 'new-register@example.com')
            self.assertTrue(any(c.isdigit() for c in password))
            self.assertTrue(any(c.isupper() for c in password))
            self.password = password
            return True
        self.mail.side_effect = mail
        result = self.request('POST', '/personnel-register', status=201, json=self.payload()).json()
        employee = result['employee']
        self.assertTrue(result['account_creation']['welcome_email_sent'])
        self.assertEqual(employee['email'], 'new-register@example.com')
        self.assertTrue(employee['linked_username'])
        self.assertIsNotNone(result['periods'][0]['account_contract_id'])
        with self.db.cursor() as cursor:
            cursor.execute('SELECT * FROM users WHERE username=%s', (employee['linked_username'],))
            user = cursor.fetchone()
        self.assertTrue(workflow.main.verify_password(self.password, user['password_hash']))
        self.assertEqual(user['manager_username'], 'fixture-manager')
        self.assertTrue(user['initial_password_pending'])
        self.assertNotIn(self.password, json.dumps(result))
        history = self.request('GET', f"/personnel-register/{employee['employee_id']}/history").text
        self.assertNotIn(self.password, history)
        self.mail.assert_called_once()

    def test_audit_failure_rolls_back_account_register_contract_roles_and_groups(self):
        tables = ['users', 'personnel_register', 'personnel_contracts', 'account_role_events', 'user_group_memberships']
        before = {table: self.count(table) for table in tables}
        self.proxy.fail_audit = True
        self.request('POST', '/personnel-register', status=500, json=self.payload())
        self.assertEqual(before, {table: self.count(table) for table in tables})
        self.mail.assert_not_called()

    def test_duplicate_email_and_unauthorized_roles_do_not_create_register(self):
        payload = self.payload()
        self.request('POST', '/personnel-register', status=201, json=payload)
        self.request('POST', '/personnel-register', status=400, json=payload)
        self.assertEqual(self.count('personnel_register'), 1)
        payload['email'] = 'other-register@example.com'
        payload['account']['role_tags'] = ['SuperAdmin']
        self.request('POST', '/personnel-register', status=403, json=payload)
        self.assertEqual(self.count('personnel_register'), 1)
        self.mail.assert_called_once()

    def test_responsable_creation_and_non_admin_denial(self):
        payload = self.payload()
        payload['account']['role_tags'] = ['Employe', 'Responsable']
        payload['account']['managed_group_ids'] = [1]
        self.request('POST', '/personnel-register', actor='fixture-manager', status=403, json=payload)
        result = self.request('POST', '/personnel-register', status=201, json=payload).json()
        with self.db.cursor() as cursor:
            cursor.execute('SELECT role_tags FROM users WHERE username=%s', (result['employee']['linked_username'],))
            self.assertIn('Responsable', json.loads(cursor.fetchone()['role_tags']))

    def test_mail_failure_keeps_both_records_and_no_account_option_sends_nothing(self):
        result = self.request('POST', '/personnel-register', status=201, json=self.payload()).json()
        self.assertFalse(result['account_creation']['welcome_email_sent'])
        self.assertEqual(self.count('personnel_register'), 1)
        self.mail.reset_mock()
        result = self.request('POST', '/personnel-register', status=201, json=registration_payload(email='optional@example.com')).json()
        self.assertIsNone(result['employee']['linked_username'])
        self.assertNotIn('account_creation', result)
        self.mail.assert_not_called()

    def test_account_only_flow_never_writes_register_or_contract(self):
        payload = {'first_name': 'External', 'last_name': 'Fixture', 'new_mail': 'external@example.com',
                   'generate_password': True, 'role_tags': ['Employe'], 'manager_id': 'fixture-manager'}
        self.request('POST', '/create-user', json=payload)
        self.assertEqual(self.count('personnel_register'), 0)
        self.assertEqual(self.count('personnel_contracts'), 0)
        self.mail.assert_called_once()
        with self.db.cursor() as cursor:
            cursor.execute('SELECT username,first_name,last_name FROM users WHERE email=%s', (payload['new_mail'],))
            account = cursor.fetchone()
        self.assertEqual(account['first_name'], 'External')
        self.assertTrue(account['username'])

    def test_employee_can_be_created_and_saved_without_manager(self):
        payload = self.payload()
        payload['account']['manager_id'] = None
        result = self.request('POST', '/personnel-register', status=201, json=payload).json()
        username = result['employee']['linked_username']
        self.request('PUT', f'/users/{username}/organization', json={
            'role_tags': ['Employe'], 'manager_id': None, 'group_ids': [], 'managed_group_ids': []})
        with self.db.cursor() as cursor:
            cursor.execute('SELECT manager_username FROM users WHERE username=%s', (username,))
            self.assertIsNone(cursor.fetchone()['manager_username'])

    def test_pure_super_admin_needs_no_register_contract_or_manager(self):
        with self.db.cursor() as cursor:
            cursor.execute("UPDATE users SET role='SuperAdmin',role_tags=%s,is_admin=1 WHERE username='fixture-admin'",
                           ('["SuperAdmin","Admin"]',))
        self.db.commit()
        self.request('POST', '/create-user', json={
            'first_name': 'System', 'last_name': 'Fixture', 'new_mail': 'system@example.com',
            'generate_password': True, 'role_tags': ['SuperAdmin', 'Admin']})
        self.assertEqual(self.count('personnel_register'), 0)
        self.assertEqual(self.count('personnel_contracts'), 0)
        with self.db.cursor() as cursor:
            cursor.execute('SELECT manager_username FROM users WHERE email=%s', ('system@example.com',))
            self.assertIsNone(cursor.fetchone()['manager_username'])

    def test_super_admin_handover_and_self_demotion_with_real_temporary_tables(self):
        with self.db.cursor() as cursor:
            cursor.execute("UPDATE users SET role='SuperAdmin',role_tags=%s,is_admin=1 WHERE username='fixture-admin'",
                           ('["SuperAdmin","Admin"]',))
        self.db.commit()
        self.request('PUT', '/users/fixture-employee/organization', json={'role_tags': ['SuperAdmin']})
        self.request('PUT', '/users/fixture-admin/organization', json={'role_tags': ['Employe'], 'manager_id': None})
        # The previous browser session loses access immediately after the commit.
        self.request('PUT', '/users/fixture-admin/organization', status=401, json={'role_tags': ['SuperAdmin']})
        with self.db.cursor() as cursor:
            cursor.execute("SELECT role_tags,session_version,remember_token_hash FROM users WHERE username='fixture-admin'")
            saved = cursor.fetchone()
        self.assertEqual(json.loads(saved['role_tags']), ['Employe'])
        self.assertEqual(saved['session_version'], 1)
        self.assertIsNone(saved['remember_token_hash'])
        last = self.client.put('/users/fixture-employee/organization',
            headers={'Authorization': 'Bearer ' + workflow.main.create_session_token('fixture-employee', 1)},
            json={'role_tags': ['Employe']})
        self.assertEqual(last.status_code, 409, last.text)
        self.mail.assert_not_called()

    def test_inactive_accounts_hidden_from_daily_list_but_admin_can_open_archive(self):
        with self.db.cursor() as cursor:
            cursor.execute("UPDATE users SET is_active=0 WHERE username='fixture-employee'")
        self.db.commit()
        for actor in ('fixture-admin', 'fixture-manager'):
            people = self.request('GET', '/list-users', actor=actor).json()['users']
            self.assertNotIn('fixture-employee', [person['username'] for person in people])
        people = self.request('GET', '/list-users?include_inactive=true').json()['users']
        self.assertIn('fixture-employee', [person['username'] for person in people])
        self.request('GET', '/list-users?include_inactive=true', actor='fixture-manager', status=403)
        self.request('GET', '/account-personnel/fixture-employee')
        detail = self.request('POST', '/personnel-register', status=201,
                              json=registration_payload(linked_username='fixture-employee')).json()
        identity = detail['employee']['employee_id']
        self.request('PATCH', f'/personnel-register/{identity}/archive', json={
            'is_archived': True, 'expected_version': detail['employee']['version']})
        active = self.request('GET', '/personnel-register?archive=active').json()['employees']
        archived = self.request('GET', '/personnel-register?archive=archived').json()['employees']
        self.assertNotIn(identity, [person['employee_id'] for person in active])
        self.assertIn(identity, [person['employee_id'] for person in archived])
        detail = self.request('GET', f'/personnel-register/{identity}').json()
        self.assertTrue(detail['employee']['is_archived'])
        self.assertEqual(detail['employee']['linked_username'], 'fixture-employee')
        self.request('GET', f'/personnel-register/{identity}/history')
        with self.db.cursor() as cursor:
            cursor.execute("UPDATE users SET is_active=1,login_deleted_at=CURRENT_TIMESTAMP WHERE username='fixture-employee'")
        self.db.commit()
        people = self.request('GET', '/list-users').json()['users']
        self.assertNotIn('fixture-employee', [person['username'] for person in people])
