"""Snapshot cryptography/calendar and real MySQL TEMPORARY-table lifecycle tests."""
from contextlib import closing
from datetime import datetime
from io import BytesIO
import json
import os
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch
import uuid

from openpyxl import load_workbook
from backend.personnel_sensitive import ensure_local_nir_key_file
from backend.snapshot_crypto import protect, reveal
from backend.snapshot_schedule import next_due
from backend.personnel_snapshots import ensure_snapshot_schema, register_snapshot_routes
from test_support.personnel_fixtures import registration_payload
import test_responsable_workflow_mysql as workflow


class ScheduleTests(unittest.TestCase):
    def test_calendar_alignment_and_no_initial_backfill(self):
        self.assertEqual(next_due(datetime(2026, 10, 29), 6), datetime(2027, 1, 1, 2))
        self.assertEqual(next_due(datetime(2026, 1, 1, 2), 6), datetime(2026, 7, 1, 1))
        self.assertEqual(next_due(datetime(2026, 10, 29), 3), datetime(2027, 1, 1, 2))
        self.assertEqual(next_due(datetime(2026, 10, 29), 1), datetime(2026, 11, 1, 2))
        self.assertEqual(next_due(datetime(2026, 7, 1, 0, 59), 6), datetime(2026, 7, 1, 1))
        with self.assertRaises(ValueError):
            next_due(datetime(2026, 1, 1), 2)


class CryptoTests(unittest.TestCase):
    def test_identity_variant_tampering_and_missing_key_fail_closed(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'keys.json'
            ensure_local_nir_key_file(path)
            with patch.dict(os.environ, {'PRESENCE_PERSONNEL_SNAPSHOT_KEYS_FILE': str(path)}):
                envelope = protect(b'PK synthetic private workbook', 'one', 'full')
                self.assertNotIn(b'private workbook', envelope)
                self.assertEqual(reveal(envelope, 'one', 'full'), b'PK synthetic private workbook')
                for identity, variant, value in [('two', 'full', envelope), ('one', 'redacted', envelope),
                                                 ('one', 'full', envelope[:-1] + bytes([envelope[-1] ^ 1]))]:
                    with self.assertRaises(Exception):
                        reveal(value, identity, variant)
            with patch.dict(os.environ, {'PRESENCE_PERSONNEL_SNAPSHOT_KEYS_FILE': ''}):
                with self.assertRaises(Exception):
                    protect(b'private', 'one', 'full')

    def test_rotation_retains_old_snapshots_when_old_key_is_retained(self):
        with tempfile.TemporaryDirectory() as directory:
            old_path, new_path = Path(directory) / 'old.json', Path(directory) / 'new.json'
            ensure_local_nir_key_file(old_path); ensure_local_nir_key_file(new_path)
            old = json.loads(old_path.read_text())
            new = json.loads(new_path.read_text())
            with patch.dict(os.environ, {'PRESENCE_PERSONNEL_SNAPSHOT_KEYS_FILE': str(old_path)}):
                envelope = protect(b'historical workbook', 'old', 'full')
            new['keys'].update(old['keys'])
            new_path.write_text(json.dumps(new))
            with patch.dict(os.environ, {'PRESENCE_PERSONNEL_SNAPSHOT_KEYS_FILE': str(new_path)}):
                self.assertEqual(reveal(envelope, 'old', 'full'), b'historical workbook')


class SnapshotCursor(workflow.WorkflowCursor):
    def execute(self, sql, params=()):
        if sql in ('SET TRANSACTION ISOLATION LEVEL REPEATABLE READ', 'START TRANSACTION WITH CONSISTENT SNAPSHOT'):
            return self.cursor.execute(sql, params)
        return super().execute(sql, params)


class SnapshotConnection(workflow.WorkflowConnection):
    def cursor(self):
        return SnapshotCursor(self)


@unittest.skipUnless(os.getenv('PRESENCE_MYSQL_TEMP_TEST') == '1', 'opt-in loopback temporary tables')
class SnapshotMySQLTests(unittest.TestCase):
    request = workflow.ResponsableWorkflowMySQLTests.request

    def setUp(self):
        workflow.ResponsableWorkflowMySQLTests.setUp(self)
        self.proxy = SnapshotConnection(self.db)
        extra = {'personnel_snapshot_settings', 'personnel_snapshots', 'personnel_snapshot_events'}
        self.stack.enter_context(patch.object(workflow.fixture, 'TABLES', workflow.fixture.TABLES | extra))
        with self.db.cursor() as raw:
            class TemporarySchema:
                def execute(_, sql, params=()):
                    return raw.execute(sql.replace('CREATE TABLE IF NOT EXISTS', 'CREATE TEMPORARY TABLE'), params)
            ensure_snapshot_schema(TemporarySchema(), datetime(2026, 10, 29))
        self.db.commit()
        directory = self.stack.enter_context(tempfile.TemporaryDirectory())
        path = Path(directory) / 'archive-keys.json'
        ensure_local_nir_key_file(path)
        self.stack.enter_context(patch.dict(os.environ, {'PRESENCE_PERSONNEL_SNAPSHOT_KEYS_FILE': str(path)}))
        self.service = register_snapshot_routes(self.client.app, get_db_connection=lambda: self.proxy,
            get_current_user=workflow.main.get_current_user, load_user=workflow.main.load_user)

    def manual(self, identity=None):
        return self.request('POST', '/personnel-snapshots', json={'request_id': identity or str(uuid.uuid4())}).json()

    def source(self):
        from test_personnel_sensitive_mysql import SYNTHETIC_NIR
        row = self.request('POST', '/personnel-register', status=201,
                           json=registration_payload(social_security_number=SYNTHETIC_NIR)).json()['employee']
        return row

    def download(self, identity, include=False, status=200, format='pdf'):
        return self.request('POST', f'/personnel-snapshots/{identity}/download', status=status,
                            json={'include_nir': include, 'confirm_sensitive': include, 'format': format})

    def test_first_start_empty_and_manual_does_not_move_schedule(self):
        before = self.request('GET', '/personnel-snapshots').json()
        self.assertEqual(before['total'], 0)
        row = self.manual()
        self.assertEqual(row['status'], 'ready')
        self.assertEqual(row['row_count'], 0)
        after = self.request('GET', '/personnel-snapshots').json()
        self.assertEqual(before['settings']['next_due'], after['settings']['next_due'])
        self.assertTrue(self.download(row['id']).content.startswith(b'%PDF-'))

    def test_immutable_variants_scope_periods_and_plaintext_never_stored(self):
        from test_personnel_sensitive_mysql import SYNTHETIC_NIR
        employee = self.source()
        identity = employee['employee_id']
        self.request('POST', f'/personnel-register/{identity}/periods', status=201, json={
            'expected_version': employee['version'], 'entry_date': '2025-01-01'})
        other = self.request('POST', '/personnel-register', status=201, json=registration_payload(full_name='Archived fixture')).json()['employee']
        self.request('PATCH', f"/personnel-register/{other['employee_id']}/archive", json={'expected_version': other['version'], 'is_archived': True})
        row = self.manual()
        self.assertEqual((row['status'], row['employee_count'], row['row_count']), ('ready', 1, 2))
        with self.db.cursor() as cursor:
            cursor.execute("UPDATE users SET role_tags=%s,is_admin=1 WHERE username='fixture-admin'", (json.dumps(['SuperAdmin', 'Admin']),))
        self.db.commit()
        plain = self.download(row['id'], format='xlsx').content
        full = self.download(row['id'], True, format='xlsx').content
        for content, expected in [(plain, None), (full, SYNTHETIC_NIR)]:
            workbook = load_workbook(BytesIO(content))
            self.assertEqual(workbook.active['F8'].value, expected)
            self.assertEqual(workbook.active['F9'].value, expected)
            workbook.close()
        with self.db.cursor() as cursor:
            cursor.execute('SELECT full_encrypted,redacted_encrypted FROM personnel_snapshots WHERE id=%s', (row['id'],))
            stored = cursor.fetchone()
            self.assertNotIn(SYNTHETIC_NIR.encode(), stored['full_encrypted'])
            self.assertNotEqual(stored['full_encrypted'], full)
            cursor.execute('UPDATE personnel_register SET full_name=%s WHERE employee_id=%s', ('Changed after snapshot', identity))
        self.db.commit()
        self.assertEqual(self.download(row['id'], format='xlsx').content, plain)
        self.assertEqual(self.download(row['id'], True, format='xlsx').content, full)
        self.assertNotIn('full_encrypted', self.request('GET', '/personnel-snapshots').text)

    def test_authorization_confirmation_and_cycle_version(self):
        row = self.manual()
        self.download(row['id'], format='xlsx', status=403)
        self.assertTrue(self.download(row['id']).content.startswith(b'%PDF-'))
        self.request('GET', '/personnel-snapshots', actor='fixture-manager', status=403)
        self.request('POST', '/personnel-snapshots', actor='fixture-employee', status=403, json={'request_id': str(uuid.uuid4())})
        settings = self.request('GET', '/personnel-snapshots').json()['settings']
        self.request('PUT', '/personnel-snapshots/settings', status=403, json={'interval_months': 3, 'expected_version': settings['version']})
        with self.db.cursor() as cursor:
            cursor.execute("UPDATE users SET role_tags=%s,is_admin=1 WHERE username='fixture-admin'", (json.dumps(['SuperAdmin', 'Admin']),))
        self.db.commit()
        self.request('PUT', '/personnel-snapshots/settings', json={'interval_months': 3, 'expected_version': settings['version']})
        self.request('PUT', '/personnel-snapshots/settings', status=409, json={'interval_months': 6, 'expected_version': settings['version']})
        row = self.manual()
        self.request('POST', f"/personnel-snapshots/{row['id']}/download", status=400, json={'include_nir': True})
        self.request('POST', f"/personnel-snapshots/{row['id']}/download", actor='fixture-manager', status=403, json={})

    def test_failure_retry_idempotency_and_no_replace_ready(self):
        identity = str(uuid.uuid4())
        with patch.dict(os.environ, {'PRESENCE_PERSONNEL_SNAPSHOT_KEYS_FILE': ''}):
            failed = self.manual(identity)
        self.assertEqual(failed['status'], 'failed')
        ready = self.request('POST', f'/personnel-snapshots/{identity}/retry').json()
        self.assertEqual(ready['status'], 'ready')
        original = self.download(identity).content
        self.assertEqual(self.manual(identity)['id'], identity)
        self.request('POST', f'/personnel-snapshots/{identity}/retry')
        self.assertEqual(self.download(identity).content, original)
        self.assertEqual(self.request('GET', '/personnel-snapshots').json()['total'], 1)
        self.mail.assert_not_called()

    def test_missed_periods_collapse_and_duplicate_tick_does_nothing(self):
        with self.db.cursor() as cursor:
            cursor.execute("UPDATE personnel_snapshot_settings SET next_due='2025-01-01 02:00:00' WHERE id=1")
        self.db.commit()
        row = self.service.create()
        self.assertEqual(row['source'], 'catchup')
        self.assertEqual(row['scheduled_for'], datetime(2025, 1, 1, 2))
        self.assertGreater(row['captured_at'], row['scheduled_for'])
        self.assertIsNone(self.service.create())
        listing = self.request('GET', '/personnel-snapshots').json()
        self.assertEqual(listing['total'], 1)

    def test_tampered_file_cannot_be_downloaded_or_marked_as_success(self):
        row = self.manual()
        with self.db.cursor() as cursor:
            cursor.execute('UPDATE personnel_snapshots SET redacted_encrypted=%s WHERE id=%s', (b'corrupt', row['id']))
        self.db.commit()
        self.download(row['id'], status=503)
        with self.db.cursor() as cursor:
            cursor.execute("SELECT COUNT(*) AS total FROM personnel_snapshot_events WHERE action='downloaded'")
            self.assertEqual(cursor.fetchone()['total'], 0)

    def test_other_connection_cannot_generate_while_writer_lock_is_held(self):
        import pymysql
        config = json.loads(Path('.local-test/rebuild/runtime.json').read_text(encoding='utf-8-sig'))
        with closing(pymysql.connect(host='127.0.0.1', port=3307, user='root', password=config['db_password'],
                                    database=config['database'], cursorclass=pymysql.cursors.DictCursor)) as other:
            with other.cursor() as cursor:
                cursor.execute("SELECT GET_LOCK(CONCAT('rh-snapshots:',DATABASE()),0) AS acquired")
                self.assertEqual(cursor.fetchone()['acquired'], 1)
            try:
                self.request('POST', '/personnel-snapshots', status=409, json={'request_id': str(uuid.uuid4())})
            finally:
                with other.cursor() as cursor:
                    cursor.execute("SELECT RELEASE_LOCK(CONCAT('rh-snapshots:',DATABASE()))")
        self.assertEqual(self.request('GET', '/personnel-snapshots').json()['total'], 0)

    def test_schedule_change_preserves_overdue_time_and_unknown_ids_fail(self):
        due = datetime(2025, 1, 1, 2)
        with self.db.cursor() as cursor:
            cursor.execute("UPDATE users SET role_tags=%s,is_admin=1 WHERE username='fixture-admin'", (json.dumps(['SuperAdmin', 'Admin']),))
            cursor.execute('UPDATE personnel_snapshot_settings SET next_due=%s WHERE id=1', (due,))
        self.db.commit()
        self.request('PUT', '/personnel-snapshots/settings', json={'interval_months': 3, 'expected_version': 1})
        self.assertEqual(self.request('GET', '/personnel-snapshots').json()['settings']['next_due'], due.isoformat())
        self.request('PUT', '/personnel-snapshots/settings', status=422, json={'interval_months': 2, 'expected_version': 2})
        self.download(str(uuid.uuid4()), status=404)

    def test_revoked_admin_cannot_download_and_row_limit_never_publishes_partial_data(self):
        self.source()
        with patch('backend.personnel_snapshots.MAX_EXPORT_ROWS', 0):
            failed = self.manual()
        self.assertEqual((failed['status'], failed['error_code']), ('failed', 'row_limit'))
        self.download(failed['id'], status=409)
        ready = self.manual()
        with self.db.cursor() as cursor:
            cursor.execute("UPDATE users SET is_active=0 WHERE username='fixture-admin'")
        self.db.commit()
        self.download(ready['id'], status=401)
