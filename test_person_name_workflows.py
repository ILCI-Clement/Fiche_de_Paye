"""Person labels with synthetic identities and an isolated SQLite database only."""
from contextlib import closing
from pathlib import Path
import sqlite3
import tempfile
import unittest
from unittest.mock import MagicMock

from fastapi import FastAPI, HTTPException
from fastapi.testclient import TestClient

from backend.leave_requests import historical_actor_name, register_leave_routes
from backend.leave_store import ensure_leave_schema
from backend.signature_flow import resolve_manager


class LeavePersonNameTests(unittest.TestCase):
    def setUp(self):
        directory = tempfile.TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        self.path = Path(directory.name) / 'synthetic-names.sqlite3'
        with closing(self.connect()) as db, db:
            db.executescript('''CREATE TABLE users(
                username TEXT PRIMARY KEY,email TEXT,manager_username TEXT);
                CREATE TABLE personnel_register(
                linked_username TEXT UNIQUE,full_name TEXT,
                social_security_number_encrypted TEXT);''')
            db.executemany('INSERT INTO users VALUES(?,?,?)', [
                ('opaque-member', 'member-mail@example.test', 'opaque-manager'),
                ('opaque-manager', 'manager-mail@example.test', None),
                ('opaque-outsider', 'outsider-mail@example.test', None),
                ('opaque-admin', 'admin-mail@example.test', None),
            ])
            db.executemany('INSERT INTO personnel_register VALUES(?,?,?)', [
                ('opaque-member', 'Synthetic Employee', 'private-nir-sentinel'),
                ('opaque-manager', 'Synthetic Manager', None),
                ('opaque-outsider', 'Synthetic Outsider', None),
            ])
        ensure_leave_schema(self.connect)
        self.actor = {'username': 'opaque-member', 'role_tags': ['Employe']}
        app = FastAPI()
        register_leave_routes(app, get_db_connection=self.connect, get_current_user=lambda: self.actor)
        self.client = TestClient(app)
        self.addCleanup(self.client.close)

    def connect(self):
        db = sqlite3.connect(self.path)
        db.row_factory = sqlite3.Row
        return db

    def create(self):
        return self.client.post('/leave-requests', json={
            'start_date': '2026-10-20', 'end_date': '2026-10-20',
            'start_slot': 'morning', 'end_slot': 'morning', 'note': '',
        })

    def test_session_and_request_use_linked_full_names(self):
        response = self.client.get('/leave/session')
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.json()['name'], 'Synthetic Employee')
        self.assertEqual(response.json()['approver_name'], 'Synthetic Manager')
        request = self.create()
        self.assertEqual(request.status_code, 201, request.text)
        self.assertEqual(request.json()['employee_name'], 'Synthetic Employee')
        self.assertEqual(request.json()['approver_name'], 'Synthetic Manager')

    def test_missing_or_blank_register_name_uses_email_prefix(self):
        with closing(self.connect()) as db, db:
            db.execute("UPDATE personnel_register SET full_name='  ' WHERE linked_username='opaque-member'")
            db.execute("DELETE FROM personnel_register WHERE linked_username='opaque-manager'")
        response = self.client.get('/leave/session').json()
        self.assertEqual(response['name'], 'member-mail')
        self.assertEqual(response['approver_name'], 'manager-mail')

    def test_unlinked_register_name_does_not_establish_identity(self):
        with closing(self.connect()) as db, db:
            db.execute("UPDATE personnel_register SET linked_username=NULL WHERE linked_username='opaque-member'")
        self.assertEqual(self.client.get('/leave/session').json()['name'], 'member-mail')
        self.assertEqual(self.create().json()['employee_name'], 'member-mail')

    def test_duplicate_display_names_do_not_grant_read_or_approval(self):
        identifier = self.create().json()['id']
        with closing(self.connect()) as db, db:
            db.execute("UPDATE personnel_register SET full_name='Synthetic Manager' WHERE linked_username='opaque-outsider'")
        self.actor = {'username': 'opaque-outsider', 'role_tags': ['Responsable']}
        self.assertEqual(self.client.get(f'/leave-requests/{identifier}').status_code, 403)
        self.assertEqual(self.client.get('/leave-requests?scope=team').json()['total'], 0)
        self.assertEqual(self.client.post(f'/leave-requests/{identifier}/decision',
                         json={'decision': 'approved', 'comment': ''}).status_code, 403)

    def test_name_change_refreshes_projection_without_rewriting_history(self):
        identifier = self.create().json()['id']
        with closing(self.connect()) as db, db:
            before = [tuple(row) for row in db.execute('SELECT * FROM leave_events')]
            saved_name = db.execute('SELECT employee_name FROM leave_requests').fetchone()[0]
            db.execute("UPDATE personnel_register SET full_name='Synthetic Renamed' WHERE linked_username='opaque-member'")
        response = self.client.get(f'/leave-requests/{identifier}')
        self.assertEqual(response.json()['request']['employee_name'], 'Synthetic Renamed')
        self.assertEqual(response.json()['history'][0]['actor_name'], 'Synthetic Renamed')
        with closing(self.connect()) as db:
            self.assertEqual([tuple(row) for row in db.execute('SELECT * FROM leave_events')], before)
            self.assertEqual(db.execute('SELECT employee_name FROM leave_requests').fetchone()[0], saved_name)

    def test_request_dto_and_history_do_not_contain_sensitive_register_fields(self):
        identifier = self.create().json()['id']
        for response in [self.client.get('/leave/session'), self.client.get('/leave-requests'),
                         self.client.get(f'/leave-requests/{identifier}')]:
            self.assertEqual(response.status_code, 200)
            self.assertNotIn('private-nir-sentinel', response.text)
            self.assertNotIn('social_security', response.text)

    def test_login_rename_keeps_recorded_actor_name_and_identity_unchanged(self):
        identifier = self.create().json()['id']
        with closing(self.connect()) as db, db:
            before = [tuple(row) for row in db.execute('SELECT * FROM leave_events')]
            db.execute("UPDATE users SET username='renamed-member' WHERE username='opaque-member'")
            db.execute("UPDATE personnel_register SET linked_username='renamed-member' WHERE linked_username='opaque-member'")
            db.execute("UPDATE leave_requests SET employee_username='renamed-member' WHERE employee_username='opaque-member'")
        self.actor['username'] = 'renamed-member'
        response = self.client.get(f'/leave-requests/{identifier}')
        self.assertEqual(response.status_code, 200)
        event = response.json()['history'][0]
        self.assertEqual(event['actor_name'], 'Synthetic Employee')
        self.assertEqual(event['actor_username'], 'opaque-member')
        with closing(self.connect()) as db:
            self.assertEqual([tuple(row) for row in db.execute('SELECT * FROM leave_events')], before)

    def test_current_manager_identity_controls_approval_after_name_change(self):
        identifier = self.create().json()['id']
        with closing(self.connect()) as db, db:
            db.execute("UPDATE users SET manager_username='opaque-outsider' WHERE username='opaque-member'")
            db.execute("UPDATE personnel_register SET full_name='Synthetic Manager' WHERE linked_username='opaque-outsider'")
        self.actor = {'username': 'opaque-manager', 'role_tags': ['Responsable']}
        self.assertEqual(self.client.post(f'/leave-requests/{identifier}/decision',
                         json={'decision': 'approved', 'comment': ''}).status_code, 403)
        self.actor = {'username': 'opaque-outsider', 'role_tags': ['Responsable']}
        self.assertEqual(self.client.post(f'/leave-requests/{identifier}/decision',
                         json={'decision': 'approved', 'comment': ''}).status_code, 200)


class SignaturePersonNameTests(unittest.TestCase):
    def setUp(self):
        self.db = MagicMock()
        self.cursor = self.db.cursor.return_value.__enter__.return_value
        self.cursor.fetchall.return_value = [{'username': 'opaque-member'}]
        self.employee = {'username': 'opaque-member', 'email': 'member-mail@example.test',
                         'full_name': 'Synthetic Employee', 'manager_username': 'opaque-manager'}
        self.manager = {'username': 'opaque-manager', 'email': 'manager-mail@example.test',
                        'full_name': 'Synthetic Manager', 'role_tags': ['Responsable']}
        self.access = MagicMock(side_effect=lambda *_args, **_kwargs: self.employee)
        self.load = MagicMock(side_effect=lambda *_args: self.manager)

    def resolve(self):
        return resolve_manager('member-mail@example.test', {'username': 'actor'},
                               lambda: self.db, self.load, self.access)

    def test_provider_labels_use_authoritative_names(self):
        self.assertEqual(self.resolve(), {'name': 'Synthetic Manager',
            'email': 'manager-mail@example.test', 'employee_name': 'Synthetic Employee'})
        self.access.assert_called_once_with(self.cursor, {'username': 'actor'}, 'opaque-member', edit=False)
        self.load.assert_called_once_with(self.cursor, 'opaque-manager')
        self.db.close.assert_called_once()

    def test_missing_full_name_uses_email_prefix_without_username(self):
        self.employee.pop('full_name')
        self.manager.pop('full_name')
        result = self.resolve()
        self.assertEqual(result['name'], 'manager-mail')
        self.assertEqual(result['employee_name'], 'member-mail')
        self.assertNotIn('opaque-', repr(result))

    def test_duplicate_employee_email_remains_blocked(self):
        self.cursor.fetchall.return_value = [{'username': 'opaque-member'}, {'username': 'opaque-other'}]
        with self.assertRaises(HTTPException) as error:
            self.resolve()
        self.assertEqual(error.exception.status_code, 400)
        self.access.assert_not_called()
        self.load.assert_not_called()

    def test_matching_name_cannot_bypass_target_authorization(self):
        self.employee['full_name'] = self.manager['full_name']
        self.access.side_effect = HTTPException(403, 'Outside scope')
        with self.assertRaises(HTTPException) as error:
            self.resolve()
        self.assertEqual(error.exception.status_code, 403)
        self.load.assert_not_called()

    def test_deactivated_employee_is_not_a_valid_signer(self):
        self.employee['is_active'] = False
        with self.assertRaises(HTTPException) as error:
            self.resolve()
        self.assertEqual(error.exception.status_code, 400)
        self.load.assert_not_called()


class HistoricalActorNameTests(unittest.TestCase):
    def test_current_account_label_has_priority_over_snapshot(self):
        event = {'actor_username': 'opaque-id', 'actor_name': 'Previous Name'}
        self.assertEqual(historical_actor_name(event, {'opaque-id': 'Current Name'}), 'Current Name')

    def test_missing_account_uses_recorded_person_name(self):
        event = {'actor_username': 'opaque-id', 'actor_name': '  Recorded Person  '}
        self.assertEqual(historical_actor_name(event, {}), 'Recorded Person')

    def test_empty_and_technical_id_snapshots_never_expose_account_name(self):
        for snapshot in [None, '', '   ', 'opaque-id', '  opaque-id  ']:
            with self.subTest(snapshot=snapshot):
                self.assertEqual(historical_actor_name({'actor_username': 'opaque-id',
                    'actor_name': snapshot}, {}), 'Ancien compte')

    def test_whole_email_snapshot_is_only_rendered_as_local_part(self):
        event = {'actor_username': 'opaque-id', 'actor_name': 'recorded.person@example.com'}
        self.assertEqual(historical_actor_name(event, {}), 'recorded.person')

    def test_invalid_or_wrapped_email_snapshots_are_not_displayed(self):
        for snapshot in ['person@@example.com', 'Person <person@example.com>', '<person@example.com>',
                         'person@example.com>', 'person@', '@example.com', 'person @example.com',
                         'person@example.com extra', 'person@example..com']:
            with self.subTest(snapshot=snapshot):
                self.assertEqual(historical_actor_name({'actor_username': 'opaque-id',
                    'actor_name': snapshot}, {}), 'Ancien compte')


if __name__ == '__main__':
    unittest.main()
