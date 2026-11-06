"""Real leave HTTP transitions and calendar persistence in isolated synthetic SQLite."""
import copy
import json
import sqlite3
from test_support.sqlite_fixtures import connect_fixture
import tempfile
import unittest
from contextlib import closing
from datetime import date
from pathlib import Path
from unittest.mock import patch

from fastapi import FastAPI, HTTPException
from fastapi.testclient import TestClient
from calendar_view import apply_calendar_status, resolve_cell
from backend.attendance_sheet import new_sheet
from backend.leave_attendance import LeaveAttendance, PAID
from backend.leave_requests import register_leave_routes
from backend.leave_store import LeaveStore, ensure_leave_schema


class LeaveAttendanceTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.path = Path(self.temp.name) / 'synthetic.sqlite'
        with closing(self.connect()) as db, db:
            db.executescript('''CREATE TABLE users(username TEXT PRIMARY KEY,email TEXT,manager_username TEXT);
              INSERT INTO users VALUES('member','member@example.com','boss'),('boss','boss@example.com',NULL),
                ('other','other@example.com',NULL),('admin','admin@example.com',NULL);
              CREATE TABLE Presence(user_id TEXT PRIMARY KEY,form_content TEXT);
              CREATE TABLE personnel_contracts(id INTEGER PRIMARY KEY,username TEXT,employee_type TEXT,
                start_date TEXT,end_date TEXT,is_cdi INTEGER,is_archived INTEGER);
              INSERT INTO personnel_contracts VALUES(1,'member','salarie','2026-01-01',NULL,1,0);
              CREATE TABLE signed_fiche_documents(employee_username TEXT,fiche_year INT,fiche_month INT,mode TEXT,status TEXT);
            ''')
        ensure_leave_schema(self.connect)
        self.actor = {'username': 'member', 'role_tags': ['Employe']}
        self.service = LeaveAttendance()
        app = FastAPI()
        register_leave_routes(app, get_db_connection=self.connect, get_current_user=lambda: self.actor, attendance=self.service)
        self.client = TestClient(app)
        self.addCleanup(self.client.close)

    def connect(self):
        db = connect_fixture(self.path)
        db.row_factory = sqlite3.Row
        return db

    def actor_as(self, username):
        self.actor = {'username': username, 'role_tags': ['Admin' if username == 'admin' else 'Responsable' if username in ('boss', 'other') else 'Employe']}

    def create(self, **extra):
        self.actor_as('member')
        response = self.client.post('/leave-requests', json={'start_date': '2026-10-30', 'end_date': '2026-10-30', **extra})
        self.assertEqual(response.status_code, 201, response.text)
        self.actor_as('boss')
        return response.json()['id']

    def decide(self, identifier, decision='approved', token=None, cancel=False):
        return self.client.post(f"/leave-requests/{identifier}/{'cancellation-decision' if cancel else 'decision'}",
                                json={'decision': decision, 'attendance_confirmation': token})

    def preview(self, identifier, operation='approve'):
        response = self.client.get(f'/leave-requests/{identifier}/attendance-preview?operation={operation}')
        self.assertEqual(response.status_code, 200, response.text)
        return response.json()

    def config(self, owner='boss'):
        with closing(self.connect()) as db:
            row = db.execute('SELECT form_content FROM Presence WHERE user_id=?', (owner,)).fetchone()
        return json.loads(row[0]) if row else {'employes_data': []}

    def save(self, config, owner='boss'):
        with closing(self.connect()) as db, db:
            db.execute('INSERT OR REPLACE INTO Presence VALUES(?,?)', (owner, json.dumps(config)))

    def sheet(self, month=10, owner='boss'):
        with closing(self.connect()) as db:
            contract = dict(db.execute('SELECT * FROM personnel_contracts WHERE id=1').fetchone())
        sheet = new_sheet({'username': 'member', 'email': 'member@example.com'}, None, contract, month, 2026, [])
        self.save({'mois': month, 'annee': 2026, 'employes_data': [sheet]}, owner)
        return sheet

    def status(self, when, period='matin', owner='boss'):
        current = date.fromisoformat(when)
        sheet = next(s for s in self.config(owner)['employes_data'] if s['sheet_month'] == current.month)
        return resolve_cell(sheet, current, period)[0]

    def request_cancel(self, identifier):
        self.actor_as('member')
        response = self.client.post(f'/leave-requests/{identifier}/cancellation')
        self.assertEqual(response.status_code, 200, response.text)
        self.actor_as('boss')

    def signed(self, month=10):
        with closing(self.connect()) as db, db:
            db.execute("INSERT INTO signed_fiche_documents VALUES('member',2026,?,'production','ready')", (month,))

    def test_submission_and_preview_are_read_only_then_approval_creates_cross_month_half_days(self):
        identifier = self.create(end_date='2026-11-03', start_slot='afternoon', end_slot='morning')
        self.assertEqual(self.config()['employes_data'], [])
        preview = self.preview(identifier)
        self.assertEqual(preview['created_sheets'], 2)
        self.assertEqual(self.config()['employes_data'], [])
        self.assertEqual(self.decide(identifier, token=preview['fingerprint']).status_code, 200)
        self.assertEqual(self.status('2026-10-30'), 'Travail')
        self.assertEqual(self.status('2026-10-30', 'aprem'), PAID)
        self.assertEqual(self.status('2026-11-01'), 'Férié')
        self.assertEqual(self.status('2026-11-02'), PAID)
        self.assertEqual(self.status('2026-11-03'), PAID)
        self.assertEqual(self.status('2026-11-03', 'aprem'), 'Travail')

    def test_updates_all_saved_copies_by_account_and_preserves_other_people(self):
        self.sheet(owner='boss')
        other = self.sheet(owner='admin')
        config = self.config('admin')
        untouched = {**copy.deepcopy(other), 'id': 'unrelated', 'account_username': 'other'}
        config['employes_data'].append(untouched)
        self.save(config, 'admin')
        identifier = self.create()
        self.assertEqual(self.decide(identifier).status_code, 200)
        self.assertEqual(self.status('2026-10-30', owner='boss'), PAID)
        self.assertEqual(self.status('2026-10-30', owner='admin'), PAID)
        self.assertEqual(self.config('admin')['employes_data'][1], untouched)

    def test_missing_or_partially_covering_contract_blocks_every_write(self):
        with closing(self.connect()) as db, db:
            db.execute("UPDATE personnel_contracts SET end_date='2026-10-30',is_cdi=0")
        identifier = self.create(end_date='2026-11-02')
        self.assertFalse(self.preview(identifier)['can_apply'])
        self.assertEqual(self.decide(identifier).status_code, 409)
        self.assertEqual(self.config()['employes_data'], [])
        self.assertEqual(self.client.get(f'/leave-requests/{identifier}').json()['request']['status'], 'pending')

    def test_sick_leave_requires_fresh_explicit_confirmation(self):
        sheet = self.sheet()
        apply_calendar_status(sheet, date(2026, 10, 30), 'matin', 'Arrêt maladie')
        self.save({'employes_data': [sheet]})
        identifier = self.create()
        preview = self.preview(identifier)
        self.assertTrue(preview['needs_confirmation'])
        self.assertEqual(self.decide(identifier).status_code, 409)
        self.assertEqual(self.status('2026-10-30'), 'Arrêt maladie')
        self.assertEqual(self.decide(identifier, token=preview['fingerprint']).status_code, 200)
        self.request_cancel(identifier)
        self.assertEqual(self.decide(identifier, cancel=True).status_code, 200)
        self.assertEqual(self.status('2026-10-30'), 'Arrêt maladie')

    def test_stale_preview_is_rejected_without_partial_approval(self):
        self.sheet()
        identifier = self.create()
        preview = self.preview(identifier)
        config = self.config()
        apply_calendar_status(config['employes_data'][0], date(2026, 10, 30), 'matin', 'Absence')
        self.save(config)
        self.assertEqual(self.decide(identifier, token=preview['fingerprint']).status_code, 409)
        self.assertEqual(self.status('2026-10-30'), 'Absence')

    def test_signed_or_archived_working_sheet_is_not_changed(self):
        self.sheet()
        original = self.config()
        self.signed()
        identifier = self.create()
        self.assertFalse(self.preview(identifier)['can_apply'])
        self.assertEqual(self.decide(identifier).status_code, 409)
        self.assertEqual(self.config(), original)

    def test_signed_correct_paid_leave_only_links_and_cancellation_keeps_original(self):
        sheet = self.sheet()
        for period in ('matin', 'aprem'):
            apply_calendar_status(sheet, date(2026, 10, 30), period, PAID)
        self.save({'employes_data': [sheet]})
        original = self.config()
        self.signed()
        identifier = self.create()
        self.assertTrue(self.preview(identifier)['can_apply'])
        self.assertEqual(self.decide(identifier).status_code, 200)
        self.assertEqual(self.config(), original)
        self.request_cancel(identifier)
        self.assertEqual(self.decide(identifier, cancel=True).status_code, 200)
        self.assertEqual(self.config(), original)

    def test_cancel_restores_only_affected_half_days_and_preserves_new_other_edits(self):
        identifier = self.create(start_slot='morning', end_slot='morning')
        self.assertEqual(self.decide(identifier).status_code, 200)
        config = self.config()
        apply_calendar_status(config['employes_data'][0], date(2026, 10, 30), 'aprem', 'Autre', 'Training')
        self.save(config)
        self.request_cancel(identifier)
        self.assertEqual(self.decide(identifier, cancel=True).status_code, 200)
        self.assertEqual(self.status('2026-10-30'), 'Travail')
        self.assertEqual(self.status('2026-10-30', 'aprem'), 'Autre')

    def test_cancel_blocks_if_original_cell_was_modified_or_since_signed(self):
        identifier = self.create()
        self.decide(identifier)
        config = self.config()
        apply_calendar_status(config['employes_data'][0], date(2026, 10, 30), 'matin', 'Absence')
        self.save(config)
        self.request_cancel(identifier)
        self.assertFalse(self.preview(identifier, 'cancel')['can_apply'])
        self.assertEqual(self.decide(identifier, cancel=True).status_code, 409)
        self.assertEqual(self.status('2026-10-30'), 'Absence')

    def test_signing_after_approval_prevents_cancellation_mutating_source(self):
        identifier = self.create()
        self.decide(identifier)
        original = self.config()
        self.signed()
        self.request_cancel(identifier)
        self.assertEqual(self.decide(identifier, cancel=True).status_code, 409)
        self.assertEqual(self.config(), original)

    def test_overlapping_approved_leaves_restore_baseline_only_after_both_cancelled(self):
        first = self.create()
        self.decide(first)
        second = self.create()
        self.decide(second)
        self.request_cancel(first)
        self.assertEqual(self.decide(first, cancel=True).status_code, 200)
        self.assertEqual(self.status('2026-10-30'), PAID)
        self.request_cancel(second)
        self.assertEqual(self.decide(second, cancel=True).status_code, 200)
        self.assertEqual(self.status('2026-10-30'), 'Travail')

    def test_calendar_and_request_roll_back_if_history_write_fails(self):
        identifier = self.create()
        with patch.object(LeaveStore, 'event', side_effect=RuntimeError('Synthetic audit failure')):
            with self.assertRaises(RuntimeError):
                self.decide(identifier)
        self.assertEqual(self.config()['employes_data'], [])
        self.assertEqual(self.client.get(f'/leave-requests/{identifier}').json()['request']['status'], 'pending')

    def test_preview_permissions_and_rejected_request_do_not_touch_attendance(self):
        identifier = self.create()
        self.actor_as('member')
        self.assertEqual(self.client.get(f'/leave-requests/{identifier}/attendance-preview?operation=approve').status_code, 403)
        self.actor_as('other')
        self.assertEqual(self.decide(identifier).status_code, 403)
        self.actor_as('boss')
        self.assertEqual(self.decide(identifier, decision='rejected').status_code, 200)
        self.assertEqual(self.config()['employes_data'], [])

    def test_new_workspace_copy_inherits_approved_leave_and_can_restore_later(self):
        identifier = self.create()
        self.decide(identifier)
        fresh = copy.deepcopy(self.config())
        fresh['employes_data'][0]['id'] = 'later-copy'
        fresh['employes_data'][0]['vacances'] = []
        with LeaveStore(self.connect, write=True) as store:
            self.service.sync_new_sheets(store, {}, fresh, {'username': 'admin'})
        self.save(fresh, 'admin')
        self.assertEqual(self.status('2026-10-30', owner='admin'), PAID)
        self.request_cancel(identifier)
        self.assertEqual(self.decide(identifier, cancel=True).status_code, 200)
        self.assertEqual(self.status('2026-10-30', owner='admin'), 'Travail')

    def test_signed_guard_rejects_content_edits_but_allows_archival_metadata(self):
        sheet = self.sheet()
        self.signed()
        original = {'employes_data': [sheet]}
        changed = copy.deepcopy(original)
        changed['employes_data'][0]['nom'] = 'Synthetic edit'
        with LeaveStore(self.connect) as store:
            with self.assertRaises(HTTPException):
                self.service.guard_signed(store, original, changed)
            changed = copy.deepcopy(original)
            changed['employes_data'][0]['is_archived'] = True
            self.service.guard_signed(store, original, changed)

    def test_contract_extension_inherits_approved_afternoon_and_cancellation_restores_it(self):
        identifier = self.create(start_slot='afternoon', end_slot='afternoon')
        self.sheet()
        previous = self.config()
        previous['employes_data'][0]['sheet_period_end'] = '2026-10-15'
        self.save(previous)
        with closing(self.connect()) as db, db:
            db.execute("UPDATE leave_requests SET status='approved' WHERE id=?", (identifier,))
        incoming = copy.deepcopy(previous)
        incoming['employes_data'][0]['sheet_period_end'] = '2026-10-31'
        with LeaveStore(self.connect, write=True) as store:
            self.service.sync_new_sheets(store, previous, incoming, {'username': 'boss'})
            store.execute('UPDATE Presence SET form_content=%s WHERE user_id=%s', (json.dumps(incoming), 'boss'))
            store.db.commit()
        self.assertEqual(self.status('2026-10-30', 'matin'), 'Travail')
        self.assertEqual(self.status('2026-10-30', 'aprem'), PAID)
        self.request_cancel(identifier)
        self.assertEqual(self.decide(identifier, cancel=True).status_code, 200)
        self.assertEqual(self.status('2026-10-30', 'aprem'), 'Travail')


if __name__ == '__main__':
    unittest.main()
