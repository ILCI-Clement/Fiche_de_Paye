"""Exercise actual scoped SELECTs on an isolated database, never production data."""
from collections import Counter
import sqlite3
import unittest
from unittest.mock import Mock

from fastapi import FastAPI
from fastapi.testclient import TestClient
from backend.work_center import register_work_center, person_readiness, mailbox


class Cursor:
    def __init__(self, db):
        self.db = db
        self.cursor = db.connection.cursor()
    def __enter__(self): return self
    def __exit__(self, *_): self.cursor.close()
    def execute(self, query, params=()):
        self.db.queries.append(query)
        assert query.lstrip().startswith('SELECT'), 'This endpoint must remain read-only'
        return self.cursor.execute(query.replace('%s', '?').replace(' AS BINARY)', ' AS BLOB)'), tuple(params))
    def fetchall(self): return [dict(row) for row in self.cursor.fetchall()]
    def fetchone(self):
        row = self.cursor.fetchone()
        return dict(row) if row else None


class WorkCenterTests(unittest.TestCase):
    def setUp(self):
        self.connection = sqlite3.connect(':memory:', check_same_thread=False)
        self.connection.row_factory = sqlite3.Row
        self.addCleanup(self.connection.close)
        self.queries = []
        self.connection.executescript('''
            CREATE TABLE users (username TEXT PRIMARY KEY, manager_username TEXT);
            CREATE TABLE personnel_contracts (id INTEGER PRIMARY KEY,username TEXT,start_date TEXT,end_date TEXT,is_archived INT);
            CREATE TABLE personnel_register (employee_id INTEGER PRIMARY KEY,linked_username TEXT,full_name TEXT,
                first_name TEXT,last_name TEXT,is_archived INT,nationality TEXT,sex TEXT,address TEXT,
                birth_date TEXT,job_title TEXT,qualification TEXT,personnel_type TEXT);
            CREATE TABLE personnel_register_periods (period_id INTEGER PRIMARY KEY,employee_id INT,entry_date TEXT);
            CREATE TABLE leave_requests (id INTEGER PRIMARY KEY,employee_username TEXT,start_date TEXT,end_date TEXT,
                status TEXT,is_archived INT,created_at TEXT);
            CREATE TABLE signed_fiche_documents (id TEXT,employee_username TEXT,manager_username TEXT,status TEXT,
                fiche_month INT,fiche_year INT,last_error TEXT,mode TEXT,created_at TEXT);
        ''')
        self.users = {}
        for name, manager, groups, roles in (
            ('admin', None, [], ['SuperAdmin', 'Admin']),
            ('manager', None, [1], ['Responsable']), ('old-manager', None, [], ['Responsable']),
            ('own', 'manager', [1], ['Employe']), ('department', 'old-manager', [1], ['Employe']),
            ('outside', 'old-manager', [2], ['Employe']), ('inactive', 'manager', [1], ['Employe']),
        ):
            self.users[name] = {'username': name, 'manager_username': manager, 'group_ids': groups,
                'managed_group_ids': [1] if name == 'manager' else [], 'role_tags': roles,
                'full_name': f'Test {name}', 'email': f'{name}@example.com',
                'is_active': name != 'inactive', 'session_version': 0}
            self.connection.execute('INSERT INTO users VALUES (?,?)', (name, manager))
        self.actor = dict(self.users['manager'])
        self.current = self.users['manager']
        self.connection.executemany('INSERT INTO personnel_contracts VALUES (?,?,?,?,?)', [
            (1, 'own', '2026-10-01', None, 0), (2, 'department', '2026-09-01', '2026-09-30', 0),
            (3, 'outside', '2026-01-01', None, 0)])
        self.connection.executemany('INSERT INTO leave_requests VALUES (?,?,?,?,?,?,?)', [
            (1, 'own', '2026-11-01', '2026-11-02', 'pending', 0, '2026-10-01'),
            (2, 'department', '2026-11-01', '2026-11-02', 'pending', 0, '2026-10-01'),
            (3, 'own', '2026-11-01', '2026-11-02', 'pending', 1, '2026-10-01'),
            (4, 'own', '2026-11-01', '2026-11-02', 'approved', 0, '2026-10-01')])
        self.connection.executemany('INSERT INTO signed_fiche_documents VALUES (?,?,?,?,?,?,?,?,?)', [
            ('own-new', 'own', 'manager', 'sent', 10, 2026, None, 'production', '2026-10-01'),
            ('own-old', 'own', 'old-manager', 'sent', 9, 2026, None, 'production', '2026-09-01'),
            ('unknown-manager', 'own', None, 'sent', 8, 2026, None, 'production', '2026-08-01'),
            ('department', 'department', 'old-manager', 'sent', 10, 2026, None, 'production', '2026-10-01'),
            ('archive', 'own', 'manager', 'completed', 9, 2026, 'safe error', 'production', '2026-10-01'),
            ('ready', 'own', 'manager', 'ready', 9, 2026, None, 'production', '2026-10-01'),
            ('sandbox', 'own', 'manager', 'sent', 9, 2026, None, 'sandbox', '2026-10-01')])
        self.catalog = Mock(side_effect=lambda _: list(self.users.values()))
        self.db = Mock()
        self.db.cursor.side_effect = lambda: Cursor(self)
        app = FastAPI()
        register_work_center(app, get_db_connection=lambda: self.db, get_current_user=lambda: self.actor,
            load_user=lambda *_: self.current, all_users=self.catalog,
            can_manage=lambda _c, a, u: u.get('manager_username') == a['username'] or
                bool(set(a.get('managed_group_ids', [])) & set(u.get('group_ids', []))))
        self.client = TestClient(app)
        self.addCleanup(self.client.close)

    def fetch(self, **params):
        return self.client.get('/dashboard/work-center', params={'month': 10, 'year': 2026, **params})

    def test_manager_sees_only_active_manageable_people_and_direct_leaves(self):
        response = self.fetch()
        self.assertEqual(response.status_code, 200, response.text)
        data = response.json()
        self.assertEqual({p['username'] for p in data['people']}, {'own', 'department'})
        self.assertEqual([t['id'] for t in data['tasks'] if t['kind'] == 'leave'], ['leave:1'])
        self.assertEqual(data['counts']['leave'], 1)

    def test_old_or_unknown_signing_manager_and_sandbox_are_not_inherited(self):
        data = self.fetch().json()
        ids = {t['id'] for t in data['tasks'] if t['kind'] in {'signature', 'archive'}}
        self.assertEqual(ids, {'signature:own-new', 'signature:archive'})
        archive = next(t for t in data['tasks'] if t['kind'] == 'archive')
        self.assertEqual(archive['destination']['month'], 9)
        self.assertEqual(archive['destination']['view'], 'signed')

    def test_manager_replacement_immediately_removes_old_history(self):
        self.connection.execute("UPDATE users SET manager_username='old-manager' WHERE username='own'")
        for user in self.users.values():
            if user['username'] == 'own': user['manager_username'] = 'old-manager'
        self.assertEqual(self.fetch().json()['counts']['signature'], 0)
        self.assertEqual(self.fetch().json()['counts']['archive'], 0)
        self.assertEqual(self.fetch().json()['counts']['leave'], 0)

    def test_differently_cased_account_is_not_a_current_direct_report(self):
        self.connection.execute("UPDATE users SET username='OWN' WHERE username='own'")
        data = self.fetch().json()
        self.assertFalse(any(t['kind'] in {'signature', 'archive'} for t in data['tasks']))

    def test_manager_never_queries_register_or_returns_private_data(self):
        response = self.fetch()
        self.assertEqual(response.headers['cache-control'], 'no-store')
        self.assertFalse(any('personnel_register' in q for q in self.queries))
        for forbidden in ('address', 'birth_date', 'social_security', 'password', 'register_link'):
            self.assertNotIn(forbidden, response.text)
        self.db.commit.assert_not_called()

    def test_admin_gets_unlinked_register_and_complete_history(self):
        self.actor = dict(self.users['admin']); self.current = self.users['admin']
        self.connection.execute("INSERT INTO personnel_register VALUES (10,NULL,'Synthetic Register','Synthetic','Register',0,'FR','F','PRIVATE ADDRESS','2000-01-01','Job','Q','salarie')")
        data = self.fetch().json()
        self.assertTrue(any(t['id'] == 'register:10' for t in data['tasks']))
        self.assertTrue(any(t['id'] == 'signature:own-old' for t in data['tasks']))
        self.assertNotIn('PRIVATE ADDRESS', str(data))
        self.assertNotIn('2000-01-01', str(data))

    def test_missing_register_does_not_block_attendance(self):
        self.actor = dict(self.users['admin']); self.current = self.users['admin']
        own = next(p for p in self.fetch().json()['people'] if p['username'] == 'own')
        self.assertTrue(own['attendance_ready'])
        self.assertTrue(any(i['code'] == 'register_link' for i in own['issues']))

    def test_month_changes_contract_readiness_and_department_manager_cannot_setup(self):
        people = {p['username']: p for p in self.fetch().json()['people']}
        self.assertTrue(people['own']['attendance_ready'])
        self.assertFalse(people['department']['attendance_ready'])
        contract = next(i for i in people['department']['issues'] if i['code'] == 'contract')
        self.assertIsNone(contract['destination'])
        september = {p['username']: p for p in self.fetch(month=9).json()['people']}
        self.assertTrue(september['department']['attendance_ready'])

    def test_archived_contract_does_not_make_dossier_ready(self):
        self.connection.execute('UPDATE personnel_contracts SET is_archived=1 WHERE username=?', ('own',))
        own = next(p for p in self.fetch().json()['people'] if p['username'] == 'own')
        self.assertFalse(own['attendance_ready'])

    def test_stale_disabled_deleted_or_demoted_actor_cannot_read(self):
        for update, expected in (({'is_active': False}, 401), ({'login_deleted_at': '2026-10-01'}, 401),
                                 ({'session_version': 1}, 401), ({'role_tags': ['Employe']}, 403)):
            self.current = {**self.users['manager'], **update}
            self.assertEqual(self.fetch().status_code, expected)
        self.catalog.assert_not_called()

    def test_employee_cannot_read_and_parameters_cannot_override_actor(self):
        self.actor = self.users['own']; self.current = self.users['own']
        self.assertEqual(self.fetch(actor='admin', username='outside').status_code, 403)
        self.catalog.assert_not_called()

    def test_invalid_month_and_year_rejected(self):
        for params in ({'month': 0}, {'month': 13}, {'year': 1999}, {'year': 2101}):
            self.assertEqual(self.fetch(**params).status_code, 422)

    def test_duplicate_email_and_invalid_manager_are_explicit_issues(self):
        user = self.users['own']
        emails = Counter([mailbox(user['email']), mailbox(user['email'])])
        value = person_readiness(user, {**self.users['manager'], 'is_active': False}, [],
            actor=self.actor, month=10, year=2026, emails=emails)
        self.assertFalse(value['contacts_ready'])
        self.assertTrue({'email', 'manager'} <= {i['code'] for i in value['issues']})


if __name__ == '__main__':
    unittest.main()
