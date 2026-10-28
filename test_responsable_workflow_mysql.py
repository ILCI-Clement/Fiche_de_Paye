"""Opt-in Responsable workflow on owned MySQL TEMPORARY tables only.

Uses real JWT/session checks, account creation, register linking, contracts,
attendance routes and the production save validator. SMTP is replaced by a mock.
No requests reach production and no permanent business table is mutated.
"""
import copy
import json
import os
import secrets
import unittest
from unittest.mock import Mock, patch

from fastapi import FastAPI
from fastapi.testclient import TestClient

import test_backend_helpers
import test_personnel_sensitive_mysql as fixture
from backend import main
from backend.attendance_api import register_attendance_routes
from backend.attendance_revision import config_revision
from backend.contract_history import register_contract_history
from backend.personnel_register import register_personnel_routes
from test_support.personnel_fixtures import registration_payload


class WorkflowCursor(fixture.RestrictedCursor):
    def executemany(self, sql, parameters):
        return sum(self.execute(sql, params) for params in parameters)

    def execute(self, sql, params=()):
        # SHOW TABLES cannot see connection-owned TEMPORARY tables.
        if sql in ("SHOW TABLES LIKE 'personnel_contracts'",):
            return self.cursor.execute("SELECT 'personnel_contracts' AS table_name")
        if sql.lstrip().upper().startswith('DELETE FROM USER_GROUP_MEMBERSHIPS'):
            return self.cursor.execute(sql, params)
        return super().execute(sql, params)


class WorkflowConnection(fixture.TemporaryConnection):
    def cursor(self):
        return WorkflowCursor(self)


@unittest.skipUnless(os.getenv('PRESENCE_MYSQL_TEMP_TEST') == '1', 'opt-in loopback MySQL temporary tables')
class ResponsableWorkflowMySQLTests(unittest.TestCase):
    def setUp(self):
        fixture.PersonnelSensitiveMySQLTests.setUp(self)
        self.client.close()
        extra_tables = {'Presence', 'organization_groups', 'user_group_memberships',
                        'account_role_events', 'account_username_reservations', 'contract_setup_events'}
        self.stack.enter_context(patch.object(fixture, 'TABLES', fixture.TABLES | extra_tables))
        with self.db.cursor() as cursor:
            cursor.execute('''ALTER TABLE users
                ADD password_hash VARCHAR(255) DEFAULT '', ADD is_admin BOOLEAN DEFAULT 0,
                ADD role VARCHAR(30) DEFAULT 'Employe', ADD manager_username VARCHAR(255) NULL,
                ADD created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                ADD first_name VARCHAR(100) DEFAULT 'Synthetic', ADD last_name VARCHAR(100) DEFAULT 'Fixture',
                ADD initial_password_pending BOOLEAN DEFAULT 0,
                ADD remember_token_hash CHAR(64) NULL, ADD remember_token_expires DATETIME NULL,
                ADD reset_token VARCHAR(255) NULL, ADD reset_token_expires DATETIME NULL''')
            definitions = {
                'Presence': 'user_id VARCHAR(255) PRIMARY KEY,form_content LONGTEXT',
                'organization_groups': 'id INT PRIMARY KEY,name VARCHAR(100) UNIQUE,is_active BOOLEAN DEFAULT 1',
                'user_group_memberships': 'username VARCHAR(255),group_id INT,relation VARCHAR(20)',
                'account_role_events': '''id INT AUTO_INCREMENT PRIMARY KEY,target_username VARCHAR(255),
                    actor_username VARCHAR(255),action VARCHAR(32),roles_before TEXT,roles_after TEXT''',
                'account_username_reservations': 'username VARCHAR(255) PRIMARY KEY',
                'contract_setup_events': '''id INT AUTO_INCREMENT PRIMARY KEY,contract_id BIGINT,
                    employee_username VARCHAR(255),actor_username VARCHAR(255),fiche_month INT,fiche_year INT,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP''',
            }
            for table, columns in definitions.items():
                cursor.execute(f'CREATE TEMPORARY TABLE {table} ({columns}) ENGINE=InnoDB')
            cursor.execute("INSERT INTO organization_groups(id,name) VALUES(1,'Synthetic team'),(2,'Other team')")
            cursor.execute("INSERT INTO user_group_memberships VALUES('fixture-manager',1,'manager')")
            cursor.execute("INSERT INTO user_group_memberships VALUES('fixture-employee',1,'member')")
            cursor.execute("UPDATE users SET manager_username='fixture-manager' WHERE username='fixture-employee'")
        self.db.commit()
        self.proxy = WorkflowConnection(self.db)
        self.stack.enter_context(patch.object(main, 'get_db_connection', lambda: self.proxy))
        self.mail = self.stack.enter_context(patch.object(main, 'send_account_welcome_email', Mock(return_value=False), create=True))
        app = FastAPI()
        app.add_api_route('/create-user', main.create_user, methods=['POST'])
        app.add_api_route('/list-users', main.list_users, methods=['GET'])
        app.add_api_route('/groups', main.list_groups, methods=['GET'])
        app.add_api_route('/groups/{group_id}', main.update_group, methods=['PUT'])
        app.add_api_route('/users/{target_username}/organization', main.update_user_organization, methods=['PUT'])
        register_personnel_routes(app, get_db_connection=lambda: self.proxy,
                                 get_current_user=main.get_current_user, load_user=main.load_user)
        register_contract_history(app, lambda: self.proxy, main.require_roles,
                                  main.get_current_user, main.require_target_access, load_user=main.load_user)
        register_attendance_routes(app, get_db_connection=lambda: self.proxy,
                                   get_current_user=main.get_current_user, load_user=main.load_user,
                                   can_manage=main.user_can_manage, save_config=main.save_config,
                                   save_checked=main.save_config_checked)
        # Schema is supplied by the owned fixture; do not execute app lifespan.
        self.client = TestClient(app)
        self.addCleanup(self.client.close)

    def request(self, method, path, *, actor='fixture-admin', status=200, **kwargs):
        response = self.client.request(method, path, headers={
            'Authorization': 'Bearer ' + main.create_session_token(actor)}, **kwargs)
        self.assertEqual(response.status_code, status, response.text)
        return response

    def workspace(self):
        return self.request('GET', '/attendance/workspace', actor='fixture-manager').json()

    def open_sheet(self, username, month=10, status=200):
        return self.request('POST', f'/attendance/open/{username}', actor='fixture-manager', status=status,
                            json={'month': month, 'year': 2026, 'revision': self.workspace()['revision']})

    def create_account(self, username='workflow-person', **extra):
        self.request('POST', '/create-user', json={
            'new_username': username, 'new_mail': username + '@example.com',
            'first_name': 'Synthetic', 'last_name': username,
            'new_password': secrets.token_hex(12), 'role_tags': ['Employe'],
            'employee_type': 'salarie', 'group_ids': [1], 'manager_id': 'fixture-manager', **extra})
        return username

    def test_register_only_then_link_missing_contract_then_manager_create_apply_save_reopen(self):
        registered = self.request('POST', '/personnel-register', status=201,
                                  json=registration_payload()).json()
        with self.proxy.cursor() as cursor:
            cursor.execute('SELECT COUNT(*) AS total FROM users')
            self.assertEqual(cursor.fetchone()['total'], 3, 'Registration must not create a login account')
        self.assertEqual([u['username'] for u in self.workspace()['accounts']], ['fixture-employee'])
        username = self.create_account()
        employee = registered['employee']
        self.request('PUT', f'/account-personnel/{username}/link', json={
            'employee_id': employee['employee_id'], 'target_version': employee['version'],
            'expected_employee_id': None, 'expected_version': None})
        people = self.request('GET', '/list-users', actor='fixture-manager').json()['users']
        self.assertIn(username, [u['username'] for u in people])
        missing = self.open_sheet(username, status=409)
        self.assertIn('contrat actif', missing.json()['detail'])
        self.request('POST', f'/users/{username}/contracts', json={
            'employee_type': 'salarie', 'start_date': '2026-10-01', 'end_date': '2026-11-30', 'is_cdi': False})
        linked = self.request('GET', f"/personnel-register/{employee['employee_id']}").json()
        self.assertTrue(any(p['account_contract_id'] is not None for p in linked['periods']))
        opened = self.open_sheet(username).json()
        identity = opened['active_id']
        applied = self.request('POST', f'/attendance/calendar/{identity}', actor='fixture-manager', json={
            'month': 10, 'year': 2026, 'revision': opened['revision'],
            'cells': [{'date': '2026-10-05', 'period': 'matin'}], 'status': 'Absence', 'reason': 'Synthetic test'}).json()
        self.request('POST', '/attendance/save', actor='fixture-manager', json={
            'config': applied['config'], 'revision': applied['revision']})
        reloaded = self.workspace()
        sheet = next(e for e in reloaded['config']['employes_data'] if e['id'] == identity)
        self.assertTrue(any(e['date'] == '2026-10-05' and e['matin'] for e in sheet['absences']))
        calendar = self.request('GET', f'/attendance/calendar/{identity}?month=10&year=2026',
                                actor='fixture-manager').json()
        self.assertEqual(next(d for d in calendar['days'] if d['number'] == 5)['matin'][0], 'Absence')
        self.assertEqual(self.open_sheet(username).json()['active_id'], identity)
        november = self.open_sheet(username, month=11).json()
        self.assertNotEqual(november['active_id'], identity)
        self.assertEqual(len(november['config']['employes_data']), 2)
        self.open_sheet(username, month=12, status=409)

    def test_complete_linked_register_can_generate_contract_without_second_manual_entry(self):
        username = self.create_account()
        self.request('POST', '/personnel-register', status=201, json=registration_payload(
            linked_username=username, initial_period={
                'entry_date': '2026-10-01', 'exit_date': '2026-11-30', 'contract_type': 'CDD'}))
        contracts = self.request('GET', f'/users/{username}/contracts').json()['contracts']
        self.assertEqual(len(contracts), 1)
        self.assertEqual(contracts[0]['start_date'], '2026-10-01')
        self.open_sheet(username)

    def test_cdi_alternant_and_intern_contracts_allow_pure_manager_creation(self):
        for index, (kind, cdi) in enumerate([('salarie', True), ('alternance', False), ('stagiaire', False)]):
            username = self.create_account(f'workflow-kind-{index}', employee_type='stagiaire' if kind == 'stagiaire' else 'salarie')
            self.request('POST', f'/users/{username}/contracts', json={
                'employee_type': kind, 'start_date': '2026-10-01',
                'end_date': None if cdi else '2026-11-30', 'is_cdi': cdi})
            self.open_sheet(username)

    def test_scope_and_standard_role_and_identity_tampering_are_denied(self):
        username = self.create_account()
        self.request('POST', f'/users/{username}/contracts', json={
            'employee_type': 'salarie', 'start_date': '2026-10-01', 'end_date': '2026-11-30', 'is_cdi': False})
        opened = self.open_sheet(username).json()
        changed = copy.deepcopy(opened['config'])
        changed['employes_data'][0]['nom'] = 'Identity tampering'
        self.request('POST', '/attendance/save', actor='fixture-manager', status=403,
                     json={'config': changed, 'revision': opened['revision']})
        self.assertEqual(self.workspace()['revision'], opened['revision'])
        self.request('POST', '/personnel-register', actor='fixture-manager', status=403,
                     json=registration_payload())
        self.request('POST', f'/users/{username}/contracts', actor='fixture-manager', status=403,
                     json={'employee_type': 'salarie', 'start_date': '2027-01-01',
                           'end_date': '2027-12-31', 'is_cdi': False})
        self.request('GET', '/attendance/workspace', actor=username, status=403)
        self.request('POST', f'/attendance/open/{username}', actor=username, status=403,
                     json={'month': 10, 'year': 2026, 'revision': opened['revision']})
        outsider = self.create_account('workflow-outsider', group_ids=[2], manager_id='fixture-admin')
        self.assertNotIn(outsider, [u['username'] for u in self.workspace()['accounts']])
        self.open_sheet(outsider, status=403)
        with self.proxy.cursor() as cursor:
            cursor.execute('SELECT role_tags FROM users WHERE username=%s', ('fixture-manager',))
            self.assertEqual(json.loads(cursor.fetchone()['role_tags']), ['Responsable'])

    def test_admin_and_manager_have_separate_attendance_workspaces(self):
        username = self.create_account()
        self.request('POST', f'/users/{username}/contracts', json={
            'employee_type': 'salarie', 'start_date': '2026-10-01', 'end_date': '2026-11-30', 'is_cdi': False})
        admin_workspace = self.request('GET', '/attendance/workspace').json()
        admin_sheet = self.request('POST', f'/attendance/open/{username}', json={
            'month': 10, 'year': 2026, 'revision': admin_workspace['revision']}).json()
        self.assertEqual(self.workspace()['config']['employes_data'], [])
        manager_sheet = self.open_sheet(username).json()
        self.assertNotEqual(manager_sheet['active_id'], admin_sheet['active_id'])

    def test_departed_team_member_is_hidden_and_preserved_when_saving_remaining_team(self):
        username = self.create_account()
        self.request('POST', f'/users/{username}/contracts', json={
            'employee_type': 'salarie', 'start_date': '2026-10-01', 'end_date': '2026-11-30', 'is_cdi': False})
        opened = self.open_sheet(username).json()
        with self.proxy.cursor() as cursor:
            cursor.execute('UPDATE users SET manager_username=%s WHERE username=%s', ('fixture-admin', username))
            cursor.execute('UPDATE user_group_memberships SET group_id=2 WHERE username=%s', (username,))
        self.proxy.commit()
        projected = self.workspace()
        self.assertEqual(projected['config']['employes_data'], [])
        self.assertEqual(projected['hidden_sheet_count'], 1)
        self.assertNotIn(username, json.dumps(projected))
        saved = self.request('POST', '/attendance/save', actor='fixture-manager', json={
            'config': projected['config'], 'revision': projected['revision']}).json()
        self.assertEqual(saved['config']['employes_data'], [])
        self.assertEqual(saved['hidden_sheet_count'], 1)
        with self.proxy.cursor() as cursor:
            cursor.execute('SELECT form_content FROM Presence WHERE user_id=%s', ('fixture-manager',))
            self.assertEqual(json.loads(cursor.fetchone()['form_content']), opened['config'])

    def test_legacy_unlinked_sheet_no_longer_blocks_create_save_and_reopen(self):
        username = self.create_account()
        legacy = {'mois': 9, 'annee': 2026, 'employes_data': [
            {'id': 'synthetic-legacy', 'nom': 'Synthetic historical person', 'calendar_overrides': {}}]}
        with self.proxy.cursor() as cursor:
            cursor.execute('INSERT INTO Presence(user_id,form_content) VALUES(%s,%s)',
                           ('fixture-manager', json.dumps(legacy)))
        self.proxy.commit()
        visible = self.request('GET', '/list-users', actor='fixture-manager').json()['users']
        self.assertIn(username, [person['username'] for person in visible])
        projected = self.workspace()
        self.assertEqual(projected['config']['employes_data'], [])
        self.assertEqual(projected['hidden_sheet_count'], 1)
        self.assertNotIn('Synthetic historical person', json.dumps(projected))
        missing_contract = self.request('POST', f'/attendance/open/{username}', actor='fixture-manager', status=409,
                     json={'month': 10, 'year': 2026, 'revision': config_revision(legacy)})
        self.assertIn('contrat actif', missing_contract.json()['detail'])
        with self.proxy.cursor() as cursor:
            cursor.execute('SELECT form_content FROM Presence WHERE user_id=%s', ('fixture-manager',))
            self.assertEqual(json.loads(cursor.fetchone()['form_content']), legacy)
        self.request('POST', f'/users/{username}/contracts', json={
            'employee_type': 'salarie', 'start_date': '2026-10-01', 'end_date': '2026-11-30', 'is_cdi': False})
        opened = self.open_sheet(username).json()
        self.assertEqual(len(opened['config']['employes_data']), 1)
        self.assertEqual(opened['hidden_sheet_count'], 1)
        applied = self.request('POST', f"/attendance/calendar/{opened['active_id']}", actor='fixture-manager', json={
            'month': 10, 'year': 2026, 'revision': opened['revision'],
            'cells': [{'date': '2026-10-05', 'period': 'matin'}], 'status': 'Absence'}).json()
        self.request('POST', '/attendance/save', actor='fixture-manager', json={
            'config': applied['config'], 'revision': applied['revision']})
        self.assertEqual(self.open_sheet(username).json()['active_id'], opened['active_id'])
        with self.proxy.cursor() as cursor:
            cursor.execute('SELECT form_content FROM Presence WHERE user_id=%s', ('fixture-manager',))
            stored = json.loads(cursor.fetchone()['form_content'])
            self.assertEqual(stored['employes_data'][0], legacy['employes_data'][0])
            self.assertEqual(len(stored['employes_data']), 2)

    def test_hidden_record_cannot_be_injected_or_read_and_stale_history_is_not_overwritten(self):
        legacy = {'mois': 10, 'annee': 2026, 'employes_data': [
            {'id': 'protected-history', 'nom': 'Private synthetic history', 'calendar_overrides': {}}]}
        with self.proxy.cursor() as cursor:
            cursor.execute('INSERT INTO Presence(user_id,form_content) VALUES(%s,%s)',
                           ('fixture-manager', json.dumps(legacy)))
        self.proxy.commit()
        username = self.create_account()
        self.request('POST', f'/users/{username}/contracts', json={
            'employee_type': 'salarie', 'start_date': '2026-10-01', 'end_date': '2026-11-30', 'is_cdi': False})
        opened = self.open_sheet(username).json()
        injected = copy.deepcopy(opened['config'])
        forged = {**injected['employes_data'][0], 'id': 'protected-history'}
        injected['employes_data'].append(forged)
        self.request('POST', '/attendance/save', actor='fixture-manager', status=403,
                     json={'config': injected, 'revision': opened['revision']})
        self.request('GET', '/attendance/calendar/protected-history?month=10&year=2026', actor='fixture-manager', status=404)
        self.request('POST', '/attendance/export', actor='fixture-manager', status=404,
                     json={'month': 10, 'year': 2026, 'ids': ['protected-history'], 'format': 'original'})
        with self.proxy.cursor() as cursor:
            cursor.execute('SELECT form_content FROM Presence WHERE user_id=%s', ('fixture-manager',))
            stored = json.loads(cursor.fetchone()['form_content'])
            stored['employes_data'][0]['nom'] = 'Admin corrected synthetic history'
            cursor.execute('UPDATE Presence SET form_content=%s WHERE user_id=%s', (json.dumps(stored), 'fixture-manager'))
        self.proxy.commit()
        self.request('POST', '/attendance/save', actor='fixture-manager', status=409,
                     json={'config': opened['config'], 'revision': opened['revision']})
        with self.proxy.cursor() as cursor:
            cursor.execute('SELECT form_content FROM Presence WHERE user_id=%s', ('fixture-manager',))
            self.assertEqual(json.loads(cursor.fetchone()['form_content']), stored)

    def test_direct_manager_can_complete_missing_contract_with_audit_and_no_second_creation(self):
        username = self.create_account()
        missing = self.open_sheet(username, status=409)
        self.assertEqual(missing.headers['X-Attendance-Requirement'], 'contract')
        payload = {'month': 10, 'year': 2026, 'employee_type': 'salarie',
                   'start_date': '2026-10-01', 'end_date': '2026-11-30', 'is_cdi': False}
        result = self.request('POST', f'/attendance/contracts/{username}', actor='fixture-manager', json=payload).json()
        self.open_sheet(username)
        self.request('POST', f'/attendance/contracts/{username}', actor='fixture-manager', status=409, json=payload)
        with self.proxy.cursor() as cursor:
            cursor.execute('SELECT * FROM contract_setup_events')
            events = cursor.fetchall()
            self.assertEqual(len(events), 1)
            self.assertEqual(events[0]['actor_username'], 'fixture-manager')
            self.assertEqual(events[0]['employee_username'], username)
            self.assertEqual(events[0]['contract_id'], result['contract_id'])
        self.request('PATCH', f"/users/{username}/contracts/{result['contract_id']}", actor='fixture-manager', status=403, json=payload)
        self.request('PATCH', f"/users/{username}/contracts/{result['contract_id']}/archive", actor='fixture-manager', status=403,
                     json={'is_archived': True})

    def test_contract_completion_rejects_non_direct_scope_standard_roles_and_wrong_month(self):
        username = self.create_account()
        payload = {'month': 10, 'year': 2026, 'employee_type': 'salarie',
                   'start_date': '2026-10-01', 'end_date': '2026-11-30', 'is_cdi': False}
        self.request('POST', f'/attendance/contracts/{username}', actor=username, status=403, json=payload)
        self.request('POST', f'/attendance/contracts/{username}', actor='fixture-manager', status=400,
                     json={**payload, 'start_date': '2027-01-01', 'end_date': '2027-12-31'})
        self.request('POST', f'/attendance/contracts/{username}', actor='fixture-manager', status=400,
                     json={**payload, 'end_date': None})
        outsider = self.create_account('workflow-other-direct-manager', manager_id='fixture-admin')
        self.request('POST', f'/attendance/contracts/{outsider}', actor='fixture-manager', status=403, json=payload)
        with self.proxy.cursor() as cursor:
            cursor.execute('SELECT COUNT(*) AS total FROM personnel_contracts')
            self.assertEqual(cursor.fetchone()['total'], 0)
            cursor.execute('SELECT COUNT(*) AS total FROM contract_setup_events')
            self.assertEqual(cursor.fetchone()['total'], 0)

    def test_contract_completion_audit_failure_rolls_back_contract_and_account_dates(self):
        username = self.create_account()
        execute = WorkflowCursor.execute
        def fail_audit(cursor, sql, params=()):
            if sql.lstrip().startswith('INSERT INTO contract_setup_events'):
                raise RuntimeError('Synthetic completion audit failure')
            return execute(cursor, sql, params)
        with patch.object(WorkflowCursor, 'execute', fail_audit), self.assertRaises(RuntimeError):
            self.request('POST', f'/attendance/contracts/{username}', actor='fixture-manager', json={
                'month': 10, 'year': 2026, 'employee_type': 'salarie',
                'start_date': '2026-01-01', 'end_date': None, 'is_cdi': True})
        with self.proxy.cursor() as cursor:
            cursor.execute('SELECT COUNT(*) AS total FROM personnel_contracts')
            self.assertEqual(cursor.fetchone()['total'], 0)
            cursor.execute('SELECT contract_start_date,contract_end_date FROM users WHERE username=%s', (username,))
            self.assertEqual(cursor.fetchone(), {'contract_start_date': None, 'contract_end_date': None})
            cursor.execute('SELECT COUNT(*) AS total FROM contract_setup_events')
            self.assertEqual(cursor.fetchone()['total'], 0)

    def test_completion_partial_month_apply_save_reopen_sync_and_hidden_history_together(self):
        username = self.create_account()
        registered = self.request('POST', '/personnel-register', status=201, json=registration_payload(
            linked_username=username, initial_period={'entry_date': '2026-10-12'})).json()
        legacy = {'mois': 9, 'annee': 2026, 'employes_data': [
            {'id': 'unlinked-history', 'nom': 'Private synthetic legacy', 'legacy_unknown': {'keep': [1, 2]}}]}
        with self.proxy.cursor() as cursor:
            cursor.execute('INSERT INTO Presence(user_id,form_content) VALUES(%s,%s)',
                           ('fixture-manager', json.dumps(legacy)))
        self.proxy.commit()
        missing = self.open_sheet(username, status=409)
        self.assertEqual(missing.headers['X-Attendance-Requirement'], 'contract')
        payload = {'month': 10, 'year': 2026, 'employee_type': 'salarie',
                   'start_date': '2026-10-12', 'end_date': '2026-10-23', 'is_cdi': False}
        result = self.request('POST', f'/attendance/contracts/{username}', actor='fixture-manager', json=payload).json()
        opened = self.open_sheet(username).json()
        identity = opened['active_id']
        sheet, = opened['config']['employes_data']
        self.assertEqual((sheet['sheet_period_start'], sheet['sheet_period_end']), ('2026-10-12', '2026-10-23'))
        self.assertEqual(opened['hidden_sheet_count'], 1)
        calendar = self.request('GET', f'/attendance/calendar/{identity}?month=10&year=2026', actor='fixture-manager').json()
        disabled = {day['number']: day['disabled'] for day in calendar['days']}
        self.assertEqual([disabled[d] for d in (11, 12, 23, 24)], [True, False, False, True])
        applied = self.request('POST', f'/attendance/calendar/{identity}', actor='fixture-manager', json={
            'month': 10, 'year': 2026, 'revision': opened['revision'], 'status': 'Absence',
            'cells': [{'date': '2026-10-12', 'period': 'matin'}]}).json()
        self.request('POST', '/attendance/save', actor='fixture-manager', json={
            'config': applied['config'], 'revision': applied['revision']})
        self.assertEqual(self.open_sheet(username).json()['active_id'], identity)
        reloaded = self.request('GET', f'/attendance/calendar/{identity}?month=10&year=2026', actor='fixture-manager').json()
        self.assertEqual(next(day for day in reloaded['days'] if day['number'] == 12)['matin'][0], 'Absence')
        self.request('POST', f'/attendance/contracts/{username}', actor='fixture-manager', status=409, json=payload)
        linked = self.request('GET', f"/personnel-register/{registered['employee']['employee_id']}").json()
        period, = [period for period in linked['periods'] if period['account_contract_id'] == result['contract_id']]
        self.assertEqual(period['entry_date'], payload['start_date'])
        self.assertEqual(period['exit_date'], payload['end_date'])
        self.assertIsNone(period['actual_exit_date'])
        with self.proxy.cursor() as cursor:
            cursor.execute('SELECT form_content FROM Presence WHERE user_id=%s', ('fixture-manager',))
            stored = json.loads(cursor.fetchone()['form_content'])
            self.assertEqual(stored['employes_data'][0], legacy['employes_data'][0])
            self.assertEqual(len(stored['employes_data']), 2)
            cursor.execute('SELECT COUNT(*) AS total FROM personnel_contracts')
            self.assertEqual(cursor.fetchone()['total'], 1)
            cursor.execute('SELECT COUNT(*) AS total FROM contract_setup_events')
            self.assertEqual(cursor.fetchone()['total'], 1)

    def test_completion_rechecks_revoked_manager_and_rejects_inactive_targets_or_edit_payloads(self):
        username = self.create_account()
        payload = {'month': 10, 'year': 2026, 'employee_type': 'salarie',
                   'start_date': '2026-10-01', 'end_date': None, 'is_cdi': True}
        for field, value, status in [('is_active', 0, 401), ('session_version', 1, 401),
                                     ('login_deleted_at', '2026-10-01 12:00:00', 401),
                                     ('role_tags', '["Employe"]', 403)]:
            with self.subTest(field=field):
                with self.proxy.cursor() as cursor:
                    cursor.execute(f'UPDATE users SET {field}=%s WHERE username=%s', (value, 'fixture-manager'))
                self.proxy.commit()
                self.request('POST', f'/attendance/contracts/{username}', actor='fixture-manager', status=status, json=payload)
                with self.proxy.cursor() as cursor:
                    cursor.execute('''UPDATE users SET is_active=1,session_version=0,login_deleted_at=NULL,
                        role_tags=%s WHERE username=%s''', ('["Responsable"]', 'fixture-manager'))
                self.proxy.commit()
        for changes in ({'contract_id': 1}, {'username': 'fixture-admin'}, {'is_archived': False}):
            self.request('POST', f'/attendance/contracts/{username}', actor='fixture-manager', status=422,
                         json={**payload, **changes})
        with self.proxy.cursor() as cursor:
            cursor.execute('UPDATE users SET is_active=0 WHERE username=%s', (username,))
        self.proxy.commit()
        self.request('POST', f'/attendance/contracts/{username}', actor='fixture-manager', status=404, json=payload)
        with self.proxy.cursor() as cursor:
            cursor.execute('SELECT COUNT(*) AS total FROM personnel_contracts')
            self.assertEqual(cursor.fetchone()['total'], 0)
            cursor.execute('SELECT COUNT(*) AS total FROM contract_setup_events')
            self.assertEqual(cursor.fetchone()['total'], 0)

    def test_losing_team_scope_between_load_and_save_never_changes_preserved_sheet(self):
        username = self.create_account()
        self.request('POST', f'/attendance/contracts/{username}', actor='fixture-manager', json={
            'month': 10, 'year': 2026, 'employee_type': 'salarie',
            'start_date': '2026-10-01', 'end_date': None, 'is_cdi': True})
        opened = self.open_sheet(username).json()
        with self.proxy.cursor() as cursor:
            cursor.execute('UPDATE users SET manager_username=%s WHERE username=%s', ('fixture-admin', username))
            cursor.execute('UPDATE user_group_memberships SET group_id=2 WHERE username=%s', (username,))
        self.proxy.commit()
        self.request('POST', '/attendance/save', actor='fixture-manager', status=403,
                     json={'config': opened['config'], 'revision': opened['revision']})
        self.request('POST', f"/attendance/calendar/{opened['active_id']}", actor='fixture-manager', status=404,
                     json={'month': 10, 'year': 2026, 'revision': opened['revision'], 'status': 'Absence',
                           'cells': [{'date': '2026-10-05', 'period': 'matin'}]})
        with self.proxy.cursor() as cursor:
            cursor.execute('SELECT form_content FROM Presence WHERE user_id=%s', ('fixture-manager',))
            self.assertEqual(json.loads(cursor.fetchone()['form_content']), opened['config'])

    def test_unchanged_department_save_is_idempotent_and_unknown_id_stays_404(self):
        payload = {'name': 'Synthetic team', 'is_active': True}
        for _ in range(2):
            result = self.request('PUT', '/groups/1', json=payload).json()
            self.assertEqual(result, {'id': 1, **payload})
        self.request('PUT', '/groups/99999', status=404, json=payload)
        self.request('PUT', '/groups/1', actor='fixture-manager', status=403, json=payload)
        self.request('PUT', '/groups/1', status=400, json={'name': 'Other team', 'is_active': True})
        self.request('PUT', '/groups/1', json={'name': 'Renamed team', 'is_active': False})
        self.request('PUT', '/groups/1', json={'name': 'Renamed team', 'is_active': False})
        with self.proxy.cursor() as cursor:
            cursor.execute('SELECT name,is_active FROM organization_groups WHERE id=1')
            self.assertEqual(cursor.fetchone(), {'name': 'Renamed team', 'is_active': 0})

    def test_managed_departments_require_explicit_responsable_role(self):
        username = self.create_account()
        payload = {'role_tags': ['Employe'], 'employee_type': 'salarie', 'manager_id': 'fixture-admin',
                   'group_ids': [1], 'managed_group_ids': [1]}
        self.request('PUT', f'/users/{username}/organization', status=400, json=payload)
        self.request('POST', '/create-user', status=400, json={
            'new_username': 'invalid-managed-employee', 'new_mail': 'invalid-managed@example.com',
            'first_name': 'Synthetic', 'last_name': 'Fixture', 'new_password': secrets.token_hex(12), **payload})
        changed = self.request('PUT', f'/users/{username}/organization', json={**payload, 'role_tags': ['Employe', 'Responsable']})
        self.assertEqual(changed.status_code, 200)
        users = self.request('GET', '/list-users').json()['users']
        target = next(person for person in users if person['username'] == username)
        self.assertIn('Responsable', target['role_tags'])
        self.assertEqual(target['managed_group_ids'], [1])


if __name__ == '__main__':
    unittest.main()
