"""Synthetic HTTP tests for atomic direct-report assignments."""
import json
import sqlite3
import unittest

from fastapi import FastAPI
from fastapi.testclient import TestClient
from backend.management_reports import register_report_routes


class Cursor:
    def __init__(self, db): self.cursor = db.cursor()
    def __enter__(self): return self
    def __exit__(self, *args): self.cursor.close()
    def execute(self, sql, args=()):
        return self.cursor.execute(sql.replace('%s', '?').replace(' FOR UPDATE', ''), args)
    def fetchone(self):
        row = self.cursor.fetchone()
        return dict(row) if row else None
    def fetchall(self): return [dict(row) for row in self.cursor.fetchall()]


class Connection:
    def __init__(self, db): self.db = db
    def cursor(self): return Cursor(self.db)
    def commit(self): self.db.commit()
    def rollback(self): self.db.rollback()
    def close(self): pass


class ManagementReportsTests(unittest.TestCase):
    def setUp(self):
        self.db = sqlite3.connect(':memory:', check_same_thread=False)
        self.db.row_factory = sqlite3.Row
        self.addCleanup(self.db.close)
        self.db.executescript('''
            CREATE TABLE users(username TEXT PRIMARY KEY,role_tags TEXT,is_active INT,
                login_deleted_at TEXT,session_version INT,manager_username TEXT);
            CREATE TABLE personnel_register(linked_username TEXT,is_archived INT);
            CREATE TABLE management_assignment_events(id INTEGER PRIMARY KEY,
                actor_username TEXT,employee_username TEXT,before_manager TEXT,after_manager TEXT);
        ''')
        for name, roles, manager in [('admin',['Admin'],None),('super',['SuperAdmin'],None),
                ('boss',['Responsable'],None),('other',['Responsable'],None),
                ('member',['Employe'],'boss'),('transfer',['Employe'],'other'),('new',['Employe'],None)]:
            self.db.execute('INSERT INTO users VALUES(?,?,1,NULL,0,?)', (name,json.dumps(roles),manager))
        self.db.commit()
        self.actor = {'username':'admin','session_version':0}
        app = FastAPI()
        def load(cursor, name):
            cursor.execute('SELECT * FROM users WHERE username=%s', (name,))
            return cursor.fetchone()
        def cycle(cursor, employee, manager):
            visited = {employee}
            while manager:
                if manager in visited: return True
                visited.add(manager)
                manager = (load(cursor,manager) or {}).get('manager_username')
            return False
        register_report_routes(app,get_db_connection=lambda:Connection(self.db),
            get_current_user=lambda:self.actor,load_user=load,
            valid_manager=lambda user: bool(user['is_active'] and set(json.loads(user['role_tags'])) & {'Responsable','Admin','SuperAdmin'}),
            would_cycle=cycle)
        self.client = TestClient(app, raise_server_exceptions=False)
        self.addCleanup(self.client.close)

    def request(self, changes, manager='boss'):
        return self.client.put(f'/users/{manager}/reports', json={'changes':changes})

    def change(self, name, before, assigned):
        return dict(username=name, expected_manager_id=before, assigned=assigned)

    def state(self):
        return [tuple(row) for row in self.db.execute('SELECT * FROM users ORDER BY username')]

    def test_bulk_transfer_and_removal_to_no_manager_are_audited(self):
        response = self.request([self.change('member','boss',False),self.change('transfer','other',True),self.change('new',None,True)])
        self.assertEqual(response.status_code,200,response.text)
        self.assertEqual(response.json()['updated'],3)
        managers = dict(self.db.execute('SELECT username,manager_username FROM users'))
        self.assertIsNone(managers['member'])
        self.assertEqual(managers['transfer'],'boss')
        self.assertEqual(managers['new'],'boss')
        self.assertEqual(self.db.execute('SELECT COUNT(*) FROM management_assignment_events').fetchone()[0],3)
        self.assertEqual(json.loads(self.db.execute("SELECT role_tags FROM users WHERE username='transfer'").fetchone()[0]),['Employe'])

    def test_stale_assignment_rejects_entire_batch(self):
        before = self.state()
        result = self.request([self.change('new',None,True),self.change('member','other',False)])
        self.assertEqual(result.status_code,409)
        self.assertEqual(self.state(),before)
        self.assertEqual(self.db.execute('SELECT COUNT(*) FROM management_assignment_events').fetchone()[0],0)

    def test_cycle_and_audit_failure_roll_back_updates(self):
        self.db.execute("UPDATE users SET manager_username='member' WHERE username='boss'")
        self.db.commit()
        before = self.state()
        self.assertEqual(self.request([self.change('new',None,True),self.change('member','boss',True)]).status_code,400)
        self.assertEqual(self.state(),before)
        self.db.execute("UPDATE users SET manager_username=NULL WHERE username='boss'")
        self.db.execute('DROP TABLE management_assignment_events')
        self.db.commit()
        before = self.state()
        self.assertEqual(self.request([self.change('new',None,True)]).status_code,500)
        self.assertEqual(self.state(),before)

    def test_permissions_rechecked_from_live_account(self):
        for actor, status in [('boss',403),('member',403),('admin',200),('super',200)]:
            with self.subTest(actor=actor):
                self.actor={'username':actor,'session_version':0}
                self.assertEqual(self.request([self.change('member','boss',True)]).status_code,status)
        self.actor={'username':'admin','session_version':0}
        self.db.execute("UPDATE users SET session_version=1 WHERE username='admin'")
        self.db.commit()
        self.assertEqual(self.request([self.change('new',None,True)]).status_code,401)

    def test_hidden_privileged_and_inactive_candidates_are_rejected(self):
        self.assertEqual(self.request([self.change('super',None,True)]).status_code,404)
        self.assertEqual(self.request([self.change('new',None,True)],manager='super').status_code,404)
        for column, value in [('is_active',0),('login_deleted_at','2026-11-04')]:
            self.db.execute(f'UPDATE users SET {column}=? WHERE username=?',(value,'new'))
            self.db.commit()
            self.assertEqual(self.request([self.change('new',None,True)]).status_code,409)
        self.db.execute("INSERT INTO personnel_register VALUES('transfer',1)")
        self.db.commit()
        self.assertEqual(self.request([self.change('transfer','other',True)]).status_code,409)

    def test_payload_and_self_assignment_validation(self):
        change = self.change('new',None,True)
        self.assertEqual(self.request([change,change]).status_code,400)
        self.assertEqual(self.request([dict(change,assigned='yes')]).status_code,422)
        self.assertEqual(self.request([dict(change,role_tags=['Admin'])]).status_code,422)
        self.assertEqual(self.request([self.change('boss',None,True)]).status_code,400)
        self.assertEqual(self.request([change]*101).status_code,422)
