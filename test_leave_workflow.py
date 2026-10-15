"""Leave integration tests run solely in a temporary SQLite database."""
import sqlite3
import tempfile
import unittest
from contextlib import closing
from pathlib import Path
from unittest.mock import Mock
from fastapi import FastAPI, HTTPException
from fastapi.testclient import TestClient
from backend.leave_requests import register_leave_routes
from backend.leave_store import ensure_leave_schema, align_leave_identity_collations


class LeaveWorkflowTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.path = Path(self.temp.name) / 'leave.sqlite3'
        with closing(self.connect()) as db, db:
            db.execute('CREATE TABLE users(username TEXT PRIMARY KEY,email TEXT,manager_username TEXT)')
            db.executemany('INSERT INTO users VALUES(?,?,?)', [('member','member@example.com','boss'),
                ('boss','boss@example.com',None), ('other','other@example.com',None), ('admin','admin@example.com',None)])
        ensure_leave_schema(self.connect)
        self.actor = {'username': 'member', 'role_tags': ['Employe']}
        def current():
            if self.actor is None: raise HTTPException(401, 'Authentication required')
            return self.actor
        app = FastAPI()
        register_leave_routes(app, get_db_connection=self.connect, get_current_user=current)
        self.client = TestClient(app)
        self.addCleanup(self.client.close)

    def connect(self):
        db = sqlite3.connect(self.path)
        db.row_factory = sqlite3.Row
        return db

    def as_actor(self, name, role):
        self.actor = {'username': name, 'role_tags': [role]}

    def create(self, **changes):
        payload = dict(start_date='2026-09-01',end_date='2026-09-01',start_slot='morning',end_slot='morning',note='')
        payload.update(changes)
        return self.client.post('/leave-requests',json=payload)

    def decide(self, identifier, decision='approved'):
        return self.client.post(f'/leave-requests/{identifier}/decision',json={'decision':decision,'comment':'Reviewed'})

    def test_create_half_day_and_audit(self):
        response = self.create()
        self.assertEqual(response.status_code,201)
        item=response.json()
        self.assertEqual(item['status'],'pending'); self.assertEqual(item['start_slot'],'morning')
        history=self.client.get(f"/leave-requests/{item['id']}").json()['history']
        self.assertEqual([e['action'] for e in history],['submitted'])

    def test_current_manager_approves_and_history_is_kept(self):
        identifier=self.create().json()['id']
        self.as_actor('boss','Responsable')
        self.assertEqual(self.decide(identifier).status_code,200)
        self.as_actor('member','Employe')
        history=self.client.get(f'/leave-requests/{identifier}').json()['history']
        self.assertEqual([e['action'] for e in history],['submitted','decided'])
        self.assertEqual(history[-1]['actor_username'],'boss')

    def test_manager_change_takes_effect_immediately(self):
        identifier=self.create().json()['id']
        with closing(self.connect()) as db, db: db.execute("UPDATE users SET manager_username='other' WHERE username='member'")
        self.as_actor('boss','Responsable')
        self.assertEqual(self.decide(identifier).status_code,403)
        self.as_actor('other','Responsable')
        self.assertEqual(self.decide(identifier).status_code,200)

    def test_unrelated_employee_cannot_read_or_review(self):
        identifier=self.create().json()['id']
        self.as_actor('other','Employe')
        self.assertEqual(self.client.get(f'/leave-requests/{identifier}').status_code,403)
        self.assertEqual(self.decide(identifier).status_code,403)
        self.assertEqual(self.client.get('/leave-requests?scope=team').status_code,403)

    def test_withdraw_before_review_hidden_from_lists(self):
        identifier=self.create().json()['id']
        self.assertEqual(self.client.post(f'/leave-requests/{identifier}/withdraw').status_code,200)
        self.assertEqual(self.client.get('/leave-requests').json()['total'],0)
        self.assertEqual(self.client.get(f'/leave-requests/{identifier}').json()['history'][-1]['action'],'withdrawn')
        self.as_actor('boss','Responsable')
        self.assertEqual(self.decide(identifier).status_code,409)

    def test_cannot_withdraw_after_review_or_decide_twice(self):
        identifier=self.create().json()['id']
        self.as_actor('boss','Responsable')
        self.assertEqual(self.decide(identifier,'rejected').status_code,200)
        self.assertEqual(self.decide(identifier).status_code,409)
        self.as_actor('member','Employe')
        self.assertEqual(self.client.post(f'/leave-requests/{identifier}/withdraw').status_code,409)

    def test_admin_can_review_self_with_actor_recorded(self):
        self.as_actor('admin','Admin')
        identifier=self.create().json()['id']
        self.assertEqual(self.decide(identifier).status_code,200)
        self.assertEqual(self.client.get(f'/leave-requests/{identifier}').json()['history'][-1]['actor_username'],'admin')

    def test_manager_cannot_review_self(self):
        self.as_actor('boss','Responsable')
        identifier=self.create().json()['id']
        self.assertEqual(self.decide(identifier).status_code,403)

    def test_missing_manager_falls_back_to_admin(self):
        with closing(self.connect()) as db, db: db.execute("UPDATE users SET manager_username=NULL WHERE username='member'")
        identifier=self.create().json()['id']
        self.as_actor('admin','Admin')
        self.assertEqual(self.decide(identifier).status_code,200)

    def test_overlap_past_dates_and_optional_note_are_allowed(self):
        self.assertEqual(self.create().status_code,201)
        self.assertEqual(self.create().status_code,201)
        self.assertEqual(self.client.get('/leave-requests').json()['total'],2)

    def test_invalid_dates_and_single_day_slots_are_rejected(self):
        self.assertEqual(self.create(end_date='2026-08-31').status_code,422)
        self.assertEqual(self.create(end_slot='afternoon').status_code,422)

    def test_team_list_obeys_current_manager_and_pagination(self):
        self.create(); self.create()
        self.as_actor('boss','Responsable')
        response=self.client.get('/leave-requests?scope=team&page_size=1&page=2').json()
        self.assertEqual(response['total'],2); self.assertEqual(len(response['items']),1)
        self.as_actor('other','Responsable')
        self.assertEqual(self.client.get('/leave-requests?scope=team').json()['total'],0)

    def test_authentication_is_required(self):
        self.actor=None
        self.assertEqual(self.client.get('/leave-requests').status_code,401)
        self.assertEqual(self.create().status_code,401)


class LeaveCollationTests(unittest.TestCase):
    def test_alignment_changes_only_leave_identity_columns(self):
        store=Mock()
        target={'charset':'utf8mb4','collation':'utf8mb4_general_ci'}
        old={'charset':'utf8mb4','collation':'utf8mb4_unicode_ci'}
        store.one.side_effect=[target,old,old]
        align_leave_identity_collations(store)
        alterations=[c.args[0] for c in store.execute.call_args_list if c.args[0].startswith('ALTER')]
        self.assertEqual(len(alterations),2)
        self.assertIn('leave_requests',alterations[0]); self.assertIn('employee_username',alterations[0])
        self.assertIn('leave_events',alterations[1]); self.assertIn('NOT NULL',alterations[1])
        self.assertFalse(any('ALTER TABLE users' in sql for sql in alterations))

    def test_matching_schema_does_not_alter_tables(self):
        store=Mock(); store.one.return_value={'charset':'utf8mb4','collation':'utf8mb4_general_ci'}
        align_leave_identity_collations(store)
        self.assertFalse(any(c.args[0].startswith('ALTER') for c in store.execute.call_args_list))

    def test_unsafe_schema_metadata_is_rejected(self):
        store=Mock(); store.one.return_value={'charset':'utf8mb4','collation':'invalid; SQL'}
        with self.assertRaises(RuntimeError): align_leave_identity_collations(store)


if __name__=='__main__': unittest.main()
