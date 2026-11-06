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
import tempfile
from pathlib import Path
from unittest.mock import Mock, patch

from fastapi import FastAPI
from fastapi.testclient import TestClient

import test_backend_helpers
import test_personnel_sensitive_mysql as fixture
from backend import main
from backend.attendance_api import register_attendance_routes
from backend.attendance_revision import config_revision
from backend.attendance_signatures import register_signature_preflight
from backend.signature_attempts import SignatureAttempts
from backend.contract_history import register_contract_history
from backend.personnel_register import register_personnel_routes
from backend.team_directory import register_team_directory
from backend.management_reports import register_report_routes
from backend.leave_notifications import LeaveNotifications
from backend.leave_requests import register_leave_routes
from test_support.personnel_fixtures import registration_payload


class WorkflowCursor(fixture.RestrictedCursor):
    def close(self):
        self.cursor.close()

    def executemany(self, sql, parameters):
        return sum(self.execute(sql, params) for params in parameters)

    def execute(self, sql, params=()):
        if 'LEFT JOIN users m ON' in sql:
            # MySQL cannot reopen a TEMPORARY table as both employee and manager.
            self.cursor.execute('DROP TEMPORARY TABLE IF EXISTS leave_fixture_managers')
            self.cursor.execute('CREATE TEMPORARY TABLE leave_fixture_managers AS SELECT * FROM users')
            sql = sql.replace('LEFT JOIN users m ON', 'LEFT JOIN leave_fixture_managers m ON')
        # SHOW TABLES cannot see connection-owned TEMPORARY tables.
        if sql in ("SHOW TABLES LIKE 'personnel_contracts'",):
            return self.cursor.execute("SELECT 'personnel_contracts' AS table_name")
        if sql.lstrip().upper().startswith('DELETE FROM USER_GROUP_MEMBERSHIPS'):
            return self.cursor.execute(sql, params)
        return super().execute(sql, params)


class WorkflowConnection(fixture.TemporaryConnection):
    def begin(self):
        self.db.begin()

    def cursor(self):
        return WorkflowCursor(self)


@unittest.skipUnless(os.getenv('PRESENCE_MYSQL_TEMP_TEST') == '1', 'opt-in loopback MySQL temporary tables')
class ResponsableWorkflowMySQLTests(unittest.TestCase):
    def setUp(self):
        fixture.PersonnelSensitiveMySQLTests.setUp(self)
        self.client.close()
        extra_tables = {'Presence', 'organization_groups', 'user_group_memberships',
                        'account_role_events', 'account_username_reservations', 'contract_setup_events',
                        'leave_requests', 'leave_events', 'leave_attendance_cells', 'signed_fiche_documents',
                        'management_assignment_events', 'leave_notification_outbox'}
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
                'leave_notification_outbox': '''event_id BIGINT PRIMARY KEY,request_id BIGINT NOT NULL,
                    kind VARCHAR(24) NOT NULL,status VARCHAR(24) DEFAULT 'pending',attempts INT DEFAULT 0,
                    next_attempt_at VARCHAR(40),recipient_username VARCHAR(255),created_at VARCHAR(40),updated_at VARCHAR(40)''',
                'management_assignment_events': '''id BIGINT AUTO_INCREMENT PRIMARY KEY,
                    actor_username VARCHAR(255),employee_username VARCHAR(255),
                    before_manager VARCHAR(255),after_manager VARCHAR(255),
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP''',
                'Presence': 'user_id VARCHAR(255) PRIMARY KEY,form_content LONGTEXT',
                'organization_groups': 'id INT PRIMARY KEY,name VARCHAR(100) UNIQUE,is_active BOOLEAN DEFAULT 1',
                'user_group_memberships': 'username VARCHAR(255),group_id INT,relation VARCHAR(20)',
                'account_role_events': '''id INT AUTO_INCREMENT PRIMARY KEY,target_username VARCHAR(255),
                    actor_username VARCHAR(255),action VARCHAR(32),roles_before TEXT,roles_after TEXT''',
                'account_username_reservations': 'username VARCHAR(255) PRIMARY KEY',
                'contract_setup_events': '''id INT AUTO_INCREMENT PRIMARY KEY,contract_id BIGINT,
                    employee_username VARCHAR(255),actor_username VARCHAR(255),fiche_month INT,fiche_year INT,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP''',
                'leave_requests': '''id BIGINT AUTO_INCREMENT PRIMARY KEY,employee_username VARCHAR(255),
                    employee_name VARCHAR(255),start_date DATE,end_date DATE,start_slot VARCHAR(12),end_slot VARCHAR(12),
                    note TEXT,status VARCHAR(12),cancellation_status VARCHAR(12) DEFAULT '',is_archived INT DEFAULT 0,
                    created_at VARCHAR(40),updated_at VARCHAR(40)''',
                'leave_events': '''id BIGINT AUTO_INCREMENT PRIMARY KEY,request_id BIGINT,actor_username VARCHAR(255),
                    actor_name VARCHAR(255),action VARCHAR(20),previous_status VARCHAR(12),new_status VARCHAR(12),
                    comment TEXT,created_at VARCHAR(40)''',
                'leave_attendance_cells': '''id BIGINT AUTO_INCREMENT PRIMARY KEY,request_id BIGINT,
                    workspace_owner VARCHAR(255),sheet_id VARCHAR(64),cell_date DATE,cell_period VARCHAR(8),
                    before_state TEXT,after_state TEXT,changed INT,
                    UNIQUE(request_id,workspace_owner,sheet_id,cell_date,cell_period)''',
                'signed_fiche_documents': '''employee_username VARCHAR(255),fiche_year INT,fiche_month INT,
                    mode VARCHAR(32),status VARCHAR(32)''',
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
        app.add_api_route('/account', main.get_account, methods=['GET'])
        app.add_api_route('/groups', main.list_groups, methods=['GET'])
        app.add_api_route('/groups/{group_id}', main.update_group, methods=['PUT'])
        app.add_api_route('/users/{target_username}/organization', main.update_user_organization, methods=['PUT'])
        register_report_routes(app,get_db_connection=lambda:self.proxy,get_current_user=main.get_current_user,
            load_user=main.load_user,valid_manager=main.is_valid_direct_manager,
            would_cycle=main.would_create_management_cycle)
        register_personnel_routes(app, get_db_connection=lambda: self.proxy,
                                 get_current_user=main.get_current_user, load_user=main.load_user,
                                 insert_account=main.insert_account, welcome_account=main.welcome_created_account)
        register_contract_history(app, lambda: self.proxy, main.require_roles,
                                  main.get_current_user, main.require_target_access, load_user=main.load_user)
        register_attendance_routes(app, get_db_connection=lambda: self.proxy,
                                   get_current_user=main.get_current_user, load_user=main.load_user,
                                   can_manage=main.user_can_manage, save_config=main.save_config,
                                   save_checked=main.save_config_checked)
        self.notifications = LeaveNotifications(lambda: self.proxy, deliver=Mock(return_value='sent'))
        register_leave_routes(app, get_db_connection=lambda: self.proxy, get_current_user=main.get_current_user,
                              attendance=main.leave_attendance, notifications=self.notifications)
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

    def test_leave_notification_outbox_real_mysql_sends_to_current_manager_without_real_mail(self):
        # EmailStr intentionally rejects reserved .test mailboxes; delivery stays mocked.
        with self.db.cursor() as cursor:
            cursor.execute("UPDATE users SET email='fixture-manager@example.com' WHERE username='fixture-manager'")
            cursor.execute("UPDATE users SET email='fixture-employee@example.com' WHERE username='fixture-employee'")
        self.db.commit()
        self.request('POST','/leave-requests',actor='fixture-employee',status=201,json={
            'start_date':'2026-11-05','end_date':'2026-11-05','start_slot':'afternoon','end_slot':'afternoon'})
        self.notifications.deliver.assert_not_called()
        with patch.dict(os.environ,{'PRESENCE_APP_URL':'https://rh.example.com/','PRESENCE_SMTP_USER':'synthetic@example.com'}):
            self.assertTrue(self.notifications.dispatch_one())
        self.notifications.deliver.assert_called_once()
        with self.db.cursor() as cursor:
            cursor.execute('SELECT status,recipient_username FROM leave_notification_outbox')
            row=cursor.fetchone()
            self.assertEqual(row['status'],'sent')
            self.assertEqual(row['recipient_username'],'fixture-manager')
        self.assertFalse(self.notifications.dispatch_one())
        with self.db.cursor() as cursor:
            cursor.execute('SELECT id FROM leave_requests ORDER BY id DESC LIMIT 1')
            identifier=cursor.fetchone()['id']
        self.request('POST',f'/leave-requests/{identifier}/decision',actor='fixture-manager',
            json={'decision':'rejected','comment':'Synthetic decision'})
        with patch.dict(os.environ,{'PRESENCE_APP_URL':'https://rh.example.com/','PRESENCE_SMTP_USER':'synthetic@example.com'}):
            self.assertTrue(self.notifications.dispatch_one())
        message=self.notifications.deliver.call_args.args[0]
        self.assertEqual(message['To'],'fixture-employee@example.com')
        self.assertIn('refusée',message['Subject'])
        self.assertEqual(self.notifications.deliver.call_count,2)
        self.assertFalse(self.notifications.dispatch_one())

    def test_batch_reports_update_real_mysql_and_preserve_unrelated_membership(self):
        # Isolate direct supervision: department-wide authority is a separate permission.
        with self.db.cursor() as cursor:
            cursor.execute("DELETE FROM user_group_memberships WHERE username='fixture-manager'")
        self.db.commit()
        path = '/users/fixture-manager/reports'
        change = {'username':'fixture-employee','expected_manager_id':'fixture-manager','assigned':False}
        self.request('PUT',path,actor='fixture-manager',status=403,json={'changes':[change]})
        self.request('PUT',path,json={'changes':[change]})
        with self.db.cursor() as cursor:
            employee = main.load_user(cursor,'fixture-employee')
            self.assertIsNone(employee['manager_username'])
            self.assertEqual(employee['group_ids'],[1])
            self.assertFalse(main.user_can_manage(cursor,main.load_user(cursor,'fixture-manager'),employee))
        self.request('PUT',path,status=409,json={'changes':[change]})
        self.request('PUT',path,json={'changes':[dict(change,expected_manager_id=None,assigned=True)]})
        with self.db.cursor() as cursor:
            employee = main.load_user(cursor,'fixture-employee')
            self.assertEqual(employee['manager_username'],'fixture-manager')
            self.assertTrue(main.user_can_manage(cursor,main.load_user(cursor,'fixture-manager'),employee))
            cursor.execute('SELECT COUNT(*) AS total FROM management_assignment_events')
            self.assertEqual(cursor.fetchone()['total'],2)

    def test_contract_change_refreshes_existing_manager_sheet_without_losing_attendance(self):
        self.request('POST', '/users/fixture-employee/contracts', json={
            'employee_type':'salarie', 'start_date':'2025-09-15', 'end_date':'2026-09-14', 'is_cdi':False})
        opened = self.request('POST', '/attendance/open/fixture-employee', actor='fixture-manager',
            json={'month':9,'year':2026,'revision':self.workspace()['revision']}).json()
        identifier = opened['active_id']
        edited = self.request('POST', f'/attendance/calendar/{identifier}', actor='fixture-manager', json={
            'month':9,'year':2026,'revision':opened['revision'],
            'cells':[{'date':'2026-09-02','period':'matin'}], 'status':'Absence','reason':'synthetic'}).json()
        before = edited['config']['employes_data'][0]
        self.request('PATCH', f"/users/fixture-employee/contracts/{before['contract_id']}", json={
            'employee_type':'salarie','start_date':'2026-08-03','end_date':'2027-08-21','is_cdi':False})
        stale = self.workspace()
        self.assertTrue(stale['contract_refresh_required'])
        result = self.request('POST','/attendance/refresh-contracts',actor='fixture-manager',
            json={'revision':stale['revision']}).json()
        after = result['config']['employes_data'][0]
        self.assertEqual(after['id'], identifier)
        self.assertEqual(after['fdc'],'2027-08-21')
        self.assertEqual(after['ddc'],'2026-08-03')
        self.assertEqual(after['sheet_period_end'],'2026-09-30')
        self.assertEqual(after['calendar_overrides'],before['calendar_overrides'])
        self.assertEqual(len(after['contract_sync_history']),1)
        self.request('POST','/attendance/save',actor='fixture-manager',status=409,
            json={'config':edited['config'],'revision':edited['revision']})
        self.assertFalse(self.workspace()['contract_refresh_required'])
        # Later contract corrections cannot rewrite a source once signing has begun.
        with self.proxy.cursor() as cursor:
            cursor.execute("INSERT INTO signed_fiche_documents VALUES('fixture-employee',2026,9,'production','ready')")
        self.db.commit()
        self.request('PATCH',f"/users/fixture-employee/contracts/{before['contract_id']}",json={
            'employee_type':'salarie','start_date':'2026-08-03','end_date':'2028-08-21','is_cdi':False})
        signed = self.workspace()
        self.assertFalse(signed['contract_refresh_required'])
        preserved = self.request('POST','/attendance/refresh-contracts',actor='fixture-manager',
            json={'revision':signed['revision']}).json()
        self.assertEqual(preserved['config']['employes_data'][0], after)
        self.mail.assert_not_called()

    def test_approved_leave_creates_calendar_in_real_transaction_and_cancels_safely(self):
        self.request('POST', '/users/fixture-employee/contracts', json={
            'employee_type': 'salarie', 'start_date': '2026-01-01', 'end_date': None, 'is_cdi': True})
        identifier = self.request('POST', '/leave-requests', actor='fixture-employee', status=201,
            json={'start_date': '2026-10-30', 'end_date': '2026-11-03', 'start_slot': 'afternoon', 'end_slot': 'morning'}).json()['id']
        preview = self.request('GET', f'/leave-requests/{identifier}/attendance-preview?operation=approve', actor='fixture-manager').json()
        self.assertEqual(preview['created_sheets'], 2)
        self.request('POST', f'/leave-requests/{identifier}/decision', actor='fixture-manager',
                     json={'decision': 'approved', 'attendance_confirmation': preview['fingerprint']})
        from backend.leave_totals import annual_leave_totals
        def totals():
            with self.proxy.cursor() as cursor:
                cursor.execute('SELECT * FROM leave_requests WHERE employee_username=%s', ('fixture-employee',))
                return annual_leave_totals(cursor, 'fixture-employee', 2026, cursor.fetchall())
        self.assertEqual(totals()['approved_days'], 2)
        self.assertFalse(totals()['incomplete'])
        workspace = self.workspace()
        self.assertEqual(len(workspace['config']['employes_data']), 2)
        october = next(e for e in workspace['config']['employes_data'] if e['sheet_month'] == 10)
        path = f"/attendance/calendar/{october['id']}?month=10&year=2026"
        cell = self.request('GET', path, actor='fixture-manager').json()['days'][29]
        self.assertEqual(cell['matin'][0], 'Travail')
        self.assertEqual(cell['aprem'][0], 'Congé payé')
        # A stale browser cannot overwrite the synchronized calendar.
        self.request('POST', '/attendance/save', actor='fixture-manager', status=409,
                     json={'config': workspace['config'], 'revision': '0' * 64})
        self.request('POST', f'/leave-requests/{identifier}/cancellation', actor='fixture-employee')
        self.request('POST', f'/leave-requests/{identifier}/cancellation-decision', actor='fixture-manager', json={'decision': 'approved'})
        cell = self.request('GET', path, actor='fixture-manager').json()['days'][29]
        self.assertEqual(cell['aprem'][0], 'Travail')
        with self.proxy.cursor() as cursor:
            cursor.execute('SELECT COUNT(*) AS total FROM leave_attendance_cells WHERE request_id=%s', (identifier,))
            self.assertGreater(cursor.fetchone()['total'], 0)
        self.assertEqual(totals()['approved_days'], 0)
        self.mail.assert_not_called()

    def test_pending_leave_hint_does_not_change_status_and_signed_calendar_rejects_manual_save(self):
        self.request('POST', '/users/fixture-employee/contracts', json={
            'employee_type': 'salarie', 'start_date': '2026-01-01', 'end_date': None, 'is_cdi': True})
        opened = self.request('POST', '/attendance/open/fixture-employee', actor='fixture-manager',
            json={'month': 10, 'year': 2026, 'revision': self.workspace()['revision']}).json()
        self.request('POST', '/leave-requests', actor='fixture-employee', status=201,
            json={'start_date': '2026-10-30', 'end_date': '2026-10-30', 'start_slot': 'morning', 'end_slot': 'morning'})
        path = f"/attendance/calendar/{opened['active_id']}?month=10&year=2026"
        cell = self.request('GET', path, actor='fixture-manager').json()['days'][29]
        self.assertEqual(cell['matin'][0], 'Travail')
        self.assertTrue(cell['pending_leave']['matin'])
        self.assertNotIn('aprem', cell['pending_leave'])
        with self.proxy.cursor() as cursor:
            cursor.execute("INSERT INTO signed_fiche_documents VALUES('fixture-employee',2026,10,'production','ready')")
        self.db.commit()
        config = copy.deepcopy(opened['config'])
        config['employes_data'][0]['calendar_overrides'] = {'2026-10-30': {'matin': {'status': 'Absence', 'reason': ''}}}
        self.request('POST', '/attendance/save', actor='fixture-manager', status=409,
            json={'config': config, 'revision': opened['revision']})
        self.assertEqual(self.workspace()['config'], opened['config'])

    def test_archived_linked_employee_disappears_from_daily_catalogs_and_returns_on_restore(self):
        register_team_directory(self.client.app, get_db_connection=lambda: self.proxy,
                                get_current_user=main.get_current_user)
        employee = self.request('POST', '/personnel-register', status=201,
            json=registration_payload(first_name='Hidden', last_name='Person', full_name='Hidden Person',
                                      linked_username='fixture-employee')).json()['employee']
        def assert_visibility(visible):
            for actor in ('fixture-admin', 'fixture-manager'):
                names = [u['username'] for u in self.request('GET', '/list-users', actor=actor).json()['users']]
                self.assertEqual('fixture-employee' in names, visible)
            catalog = self.workspace()['accounts']
            self.assertEqual('fixture-employee' in [u['username'] for u in catalog], visible)
            directory = self.request('GET', '/organization-directory', actor='fixture-employee').json()
            self.assertEqual('fixture-employee' in [u['username'] for u in directory['users']], visible)
            team = self.request('GET', '/my-team', actor='fixture-employee').json()
            self.assertEqual(any(u['name'] == employee['full_name'] for u in team['members']), visible)
        assert_visibility(True)
        archived = self.request('PATCH', f"/personnel-register/{employee['employee_id']}/archive",
            json={'expected_version': employee['version'], 'is_archived': True}).json()['employee']
        assert_visibility(False)
        # Hiding a catalog entry must not revoke login or remove the admin's recovery path.
        own = self.request('GET', '/account', actor='fixture-employee').json()
        self.assertTrue(own['is_active'])
        history = self.request('GET', '/list-users?include_inactive=true').json()['users']
        self.assertTrue(next(u for u in history if u['username']=='fixture-employee')['personnel_archived'])
        self.request('PATCH', f"/personnel-register/{employee['employee_id']}/archive",
            json={'expected_version': archived['version'], 'is_archived': False})
        assert_visibility(True)
        with self.proxy.cursor() as cursor:
            cursor.execute("UPDATE users SET is_active=0 WHERE username='fixture-employee'")
        self.db.commit()
        names = [u['username'] for u in self.request('GET', '/organization-directory').json()['users']]
        self.assertNotIn('fixture-employee', names)

    def test_super_admin_identity_is_owner_only_across_person_catalogs(self):
        register_team_directory(self.client.app, get_db_connection=lambda: self.proxy,
                                get_current_user=main.get_current_user)
        with self.proxy.cursor() as cursor:
            cursor.execute("UPDATE users SET role='SuperAdmin',role_tags=%s,is_admin=1 WHERE username='fixture-employee'",
                           ('["SuperAdmin","Admin"]',))
        self.db.commit()
        for actor in ('fixture-admin', 'fixture-manager', 'fixture-employee'):
            expected = actor == 'fixture-employee'
            for path, key in [('/list-users', 'users'), ('/organization-directory', 'users'),
                              ('/attendance/workspace', 'accounts')]:
                result = self.request('GET', path, actor=actor).json()
                self.assertEqual('fixture-employee' in [u['username'] for u in result[key]], expected)
        self.assertEqual(self.request('GET', '/account', actor='fixture-employee').json()['username'], 'fixture-employee')

    def test_own_settings_for_pure_manager_do_not_depend_on_directory_or_register(self):
        directory = self.request('GET', '/list-users', actor='fixture-manager').json()['users']
        self.assertNotIn('fixture-manager', [u['username'] for u in directory])
        for actor in ('fixture-manager', 'fixture-employee', 'fixture-admin'):
            data = self.request('GET', '/account?username=someone-else', actor=actor).json()
            self.assertEqual(data['username'], actor)
            self.assertTrue(data['email'])
            self.assertNotIn('password_hash', data)
        self.assertEqual(self.client.get('/account').status_code, 401)
        with self.proxy.cursor() as cursor:
            cursor.execute("UPDATE users SET is_active=0 WHERE username='fixture-manager'")
        self.request('GET', '/account', actor='fixture-manager', status=401)

    def test_direct_manager_report_can_complete_contract_create_and_save_attendance(self):
        self.client.app.add_api_route('/get-config/{user_id}', main.get_config, methods=['GET'])
        self.client.app.add_api_route('/save-config/{user_id}', main.save_config, methods=['POST'])
        with self.proxy.cursor() as cursor:
            cursor.execute("UPDATE users SET role_tags=%s WHERE username='fixture-employee'", (json.dumps(['Responsable']),))
            cursor.execute("INSERT INTO users(username,email,role_tags) VALUES(%s,%s,%s)",
                           ('peer-manager', 'peer@example.test', json.dumps(['Responsable'])))
            cursor.execute("INSERT INTO user_group_memberships VALUES('peer-manager',1,'member')")
        self.db.commit()
        people = self.request('GET', '/list-users', actor='fixture-manager').json()['users']
        self.assertIn('fixture-employee', [u['username'] for u in people])
        self.assertNotIn('peer-manager', [u['username'] for u in people])
        self.request('GET', '/get-config/fixture-employee', actor='fixture-manager', status=403)
        self.request('POST', '/save-config/fixture-employee', actor='fixture-manager', status=403,
                     json={'employes_data': []})
        self.assertIn('fixture-employee', [u['username'] for u in self.workspace()['accounts']])
        self.assertNotIn('peer-manager', [u['username'] for u in self.workspace()['accounts']])
        missing = self.open_sheet('fixture-employee', status=409)
        self.assertIn('contrat', missing.json()['detail'].lower())
        contract = {'month': 10, 'year': 2026, 'employee_type': 'salarie',
                    'start_date': '2026-10-01', 'end_date': None, 'is_cdi': True}
        self.request('POST', '/attendance/contracts/peer-manager', actor='fixture-manager', status=403, json=contract)
        self.request('POST', '/attendance/contracts/fixture-employee', actor='fixture-manager', json=contract)
        opened = self.open_sheet('fixture-employee').json()
        identity = opened['active_id']
        applied = self.request('POST', f'/attendance/calendar/{identity}', actor='fixture-manager', json={
            'month': 10, 'year': 2026, 'revision': opened['revision'],
            'cells': [{'date': '2026-10-05', 'period': 'matin'}], 'status': 'Absence', 'reason': 'Synthetic hierarchy test'}).json()
        self.request('POST', '/attendance/save', actor='fixture-manager', json={
            'config': applied['config'], 'revision': applied['revision']})
        saved = next(e for e in self.workspace()['config']['employes_data'] if e['id']==identity)
        self.assertTrue(any(e['date']=='2026-10-05' and e['matin'] for e in saved['absences']))
        self.request('PUT', '/users/fixture-employee/organization', actor='fixture-manager', status=403,
                     json={'role_tags': ['Admin']})
        with self.proxy.cursor() as cursor:
            cursor.execute("UPDATE users SET manager_username='fixture-manager' WHERE username='fixture-admin'")
        self.db.commit()
        self.assertNotIn('fixture-admin', [u['username'] for u in self.workspace()['accounts']])
        self.request('POST', '/attendance/contracts/fixture-admin', actor='fixture-manager', status=403, json=contract)
    def test_pure_manager_signs_linked_legacy_sheet_without_register_or_new_contract(self):
        sheet = {'id': 'legacy-blank-contact', 'account_username': 'fixture-employee',
                 'nom': 'Historical fixture', 'ddc': '2026-01-01', 'fdc': '2026-12-31',
                 'responsable': 'Old manager', 'email_employe': '', 'legacy_marker': 'preserve'}
        stored = json.dumps({'employes_data': [sheet]})
        with self.proxy.cursor() as cursor:
            cursor.execute("UPDATE users SET email='synthetic-employee@example.com' WHERE username='fixture-employee'")
            cursor.execute('INSERT INTO Presence(user_id,form_content) VALUES(%s,%s)', ('fixture-manager', stored))
            cursor.execute("SELECT email FROM users WHERE username='fixture-employee'"); email = cursor.fetchone()['email']
            cursor.execute("SELECT email FROM users WHERE username='fixture-manager'"); boss_email = cursor.fetchone()['email']
            for table in ('personnel_register', 'personnel_contracts'):
                cursor.execute('SELECT COUNT(*) AS total FROM ' + table)
                self.assertEqual(cursor.fetchone()['total'], 0)
        # Real requests read committed fixtures on independent connections. The
        # shared TEMPORARY-table proxy must commit setup before the source lock's
        # finally/rollback releases its read-only transaction.
        self.db.commit()
        root = self.stack.enter_context(tempfile.TemporaryDirectory())
        sender = Mock(return_value={'document_id': 'synthetic-signed-request'})
        register_signature_preflight(self.client.app, get_current_user=main.get_current_user,
            get_db_connection=lambda: self.proxy, load_user=main.load_user, require_target_access=main.require_target_access,
            get_config=main.get_config, readiness=lambda: {'can_send': True, 'mode': 'sandbox-only'}, sender=sender,
            attempts=SignatureAttempts(Path(root) / 'attempts.sqlite3'))
        payload = {'ids': [sheet['id']], 'month': 10, 'year': 2026}
        item = self.request('POST', '/attendance/signature-preview', actor='fixture-manager', json=payload).json()['items'][0]
        self.assertEqual(item['error'], '')
        self.assertEqual(item['employee_email'], email)
        self.assertEqual(item['manager_email'], boss_email)
        with patch('backend.attendance_signatures.generate_fiche_file', return_value={'data': b'fixture', 'filename': 'fixture.xlsx'}) as generate:
            self.request('POST', '/attendance/signature-send', actor='fixture-manager', json={
                'id': sheet['id'], 'month': 10, 'year': 2026, 'fingerprint': item['fingerprint'], 'request_id': 'synthetic-contact-0001'})
        self.assertEqual(sender.call_args.args[0].recipient_email, email)
        self.assertEqual(generate.call_args.args[0]['email_responsable'], boss_email)
        with self.proxy.cursor() as cursor:
            cursor.execute('SELECT form_content FROM Presence WHERE user_id=%s', ('fixture-manager',))
            self.assertEqual(cursor.fetchone()['form_content'], stored)
        # Linked identity is authoritative: a different email in the old sheet cannot switch recipients.
        with self.proxy.cursor() as cursor:
            cursor.execute("UPDATE users SET is_active=0 WHERE username='fixture-employee'")
        blocked = self.request('POST', '/attendance/signature-preview', actor='fixture-manager', json=payload).json()['items'][0]
        self.assertTrue(blocked['error'])
        self.assertEqual(sender.call_count, 1)

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
