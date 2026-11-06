"""Private synthetic archive fixtures: no external signatures, mail or business writes."""
from contextlib import closing
from datetime import datetime, timedelta
import hashlib
import io
import json
import os
from pathlib import Path
import re
import sqlite3
from test_support.sqlite_fixtures import connect_fixture
import tempfile
import unittest
from unittest.mock import patch, Mock
from zipfile import ZipFile

from fastapi import FastAPI, HTTPException
from fastapi.testclient import TestClient
from pypdf import PdfWriter
from backend.signed_fiches import register_signed_fiche_routes, validate_signed_url, fetch_signed_pdf, callback_url
from backend.signed_fiche_store import rename_signed_fiche_owner


SCHEMA = '''CREATE TABLE users(username TEXT PRIMARY KEY,manager_username TEXT,role_tags TEXT,
    is_active INTEGER DEFAULT 1,session_version INTEGER DEFAULT 0,login_deleted_at TEXT);
CREATE TABLE signed_fiche_documents(id TEXT PRIMARY KEY,reference_hash TEXT UNIQUE,namespace TEXT,
    document_id TEXT UNIQUE,employee_username TEXT,sender_username TEXT,manager_username TEXT,employee_name TEXT,
    fiche_month INTEGER,fiche_year INTEGER,mode TEXT,status TEXT DEFAULT 'pending',is_archived INTEGER DEFAULT 0,signed_pdf_url TEXT,
    completed_at TEXT,archived_at TEXT,sha256 TEXT,file_bytes INTEGER,download_attempts INTEGER DEFAULT 0,
    next_attempt_at TEXT,lease_until TEXT,lease_token TEXT,last_error TEXT,created_at TEXT DEFAULT CURRENT_TIMESTAMP);
CREATE TABLE signed_fiche_events(id INTEGER PRIMARY KEY,archive_id TEXT,event TEXT,actor_username TEXT,
    created_at TEXT DEFAULT CURRENT_TIMESTAMP);
CREATE TABLE permanent_deletion_events(id INTEGER PRIMARY KEY,actor_username TEXT,resource_type TEXT,
    resource_id TEXT,metadata TEXT,created_at TEXT DEFAULT CURRENT_TIMESTAMP);'''


class Cursor:
    def __init__(self, connection): self.connection = connection; self.raw = connection.db.cursor()
    def __enter__(self): return self
    def __exit__(self, *_): self.raw.close()
    @property
    def rowcount(self): return self.raw.rowcount
    def execute(self, sql, params=()):
        if 'INSERT INTO signed_fiche_events' in sql and self.connection.fixture.fail_audit:
            raise RuntimeError('Synthetic audit failure')
        if 'FOR UPDATE' in sql and not self.connection.db.in_transaction:
            self.connection.db.execute('BEGIN IMMEDIATE')
        sql = sql.replace(' FOR UPDATE', '').replace('%s', '?').replace(' AS BINARY)', ' AS BLOB)')
        sql = re.sub(r'DATE_ADD\(UTC_TIMESTAMP\(6\),INTERVAL (\?|180) SECOND\)', r"datetime(UTC_TIMESTAMP(6), '+' || \1 || ' seconds')", sql)
        return self.raw.execute(sql, params)
    def fetchone(self):
        row = self.raw.fetchone()
        return dict(row) if row is not None else None
    def fetchall(self): return [dict(row) for row in self.raw.fetchall()]


class Connection:
    def __init__(self, fixture):
        self.fixture = fixture
        self.db = connect_fixture(fixture.path, timeout=10)
        self.db.row_factory = sqlite3.Row
        self.db.create_function('UTC_TIMESTAMP', 1, lambda _: fixture.now.isoformat(' ', timespec='microseconds'))
        self.db.create_function('IF', 3, lambda condition, yes, no: yes if condition else no)
    def cursor(self): return Cursor(self)
    def commit(self): self.db.commit()
    def rollback(self): self.db.rollback()
    def close(self): self.db.close()


class SignedFicheTests(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.root = Path(self.temporary.name)
        self.path = self.root / 'archive.sqlite3'
        self.now = datetime(2026, 10, 23, 12)
        self.fail_audit = False
        with closing(connect_fixture(self.path)) as db, db:
            db.executescript(SCHEMA)
            for name, role, manager in [('admin','Admin',None), ('super','SuperAdmin',None),
                    ('manager','Responsable',None),('former','Responsable',None),('employee','Employe','manager'),('other','Employe',None)]:
                db.execute('INSERT INTO users(username,role_tags,manager_username) VALUES(?,?,?)', (name, json.dumps([role]), manager))
        self.actor = self.identity('admin')
        self.app = FastAPI()
        self.service = register_signed_fiche_routes(self.app, get_db_connection=lambda: Connection(self),
            get_current_user=self.authenticated, load_user=self.load_user, root=self.root / 'pdfs')
        # No lifespan: workers/schema are exercised separately on a temporary MySQL connection.
        self.client = TestClient(self.app, raise_server_exceptions=False)
        self.addCleanup(self.client.close)
        self.addCleanup(patch.stopall)
        patch.dict(os.environ, {'PRESENCE_ESIGN_CALLBACK_URL':'https://example.test/api/esign/callback'}).start()
        writer = PdfWriter(); writer.add_blank_page(width=100, height=100)
        output = io.BytesIO(); writer.write(output); self.pdf = output.getvalue()

    def authenticated(self):
        if self.actor is None: raise HTTPException(401, 'Authentication required')
        return dict(self.actor)
    def load_user(self, cursor, name):
        cursor.execute('SELECT * FROM users WHERE username=%s', (name,))
        user = cursor.fetchone()
        if user: user['role_tags'] = json.loads(user['role_tags'])
        return user
    def identity(self, name):
        connection = Connection(self)
        try:
            with connection.cursor() as cursor: return self.load_user(cursor, name)
        finally: connection.close()
    def rows(self, sql='SELECT * FROM signed_fiche_documents', params=()):
        with self.service.store.transaction() as cursor:
            cursor.execute(sql, params); return cursor.fetchall()
    def prepare(self, **changes):
        values = dict(employee='employee',sender='admin',manager='manager',name='Synthetic Employee',month=9,year=2026,namespace='fixture',mode='production')
        values.update(changes)
        return self.service.prepare(**values)
    def payload(self, request, **changes):
        result = dict(reference_id=request['reference'],provider='clawshow_esign',namespace='fixture',document_id='fixture_doc',
            event='document.completed',all_signed=True,signed_pdf_url='https://esign.clawshow.ai/esign/fixture_doc/signed.pdf')
        result.update(changes); return result
    def notify(self, request, **changes): return self.client.post('/esign/callback', json=self.payload(request, **changes))
    def ready(self, **values):
        request = self.prepare(**values)
        document_id = 'fixture_' + request['id']
        self.assertEqual(self.notify(request, document_id=document_id, signed_pdf_url=f'https://esign.clawshow.ai/esign/{document_id}/signed.pdf').status_code, 204)
        row = self.service.store.claim(); self.service.save(row, self.pdf)
        return self.rows('SELECT * FROM signed_fiche_documents WHERE id=%s', (request['id'],))[0]

    def test_reference_capability_is_hashed_and_not_public(self):
        request = self.prepare()
        self.assertNotIn(request['reference'], repr(self.rows()))
        response = self.client.get('/signed-fiches?year=2026&month=9')
        self.assertEqual(response.status_code, 200)
        self.assertNotIn('reference_hash', response.text)
        self.assertNotIn(request['reference'], response.text)
        self.assertIn('no-store', response.headers['cache-control'])

    def delete(self, identifier, confirmation=None):
        return self.client.request('DELETE', f'/signed-fiches/{identifier}',
                                   json={'confirmation': confirmation or identifier})

    def test_super_admin_deletes_archived_pdf_and_metadata_with_independent_audit(self):
        row = self.ready()
        self.actor = self.identity('super')
        path = self.service.path(row)
        self.assertTrue(self.client.get('/signed-fiches?year=2026').json()['can_delete'])
        result = self.delete(row['id'])
        self.assertEqual(result.status_code, 200, result.text)
        self.assertFalse(result.json()['cleanup_pending'])
        self.assertFalse(path.exists())
        self.assertEqual(self.rows(), [])
        self.assertEqual(self.client.get('/signed-fiches?year=2026').json()['total'], 0)
        self.assertEqual(self.client.get(f"/signed-fiches/{row['id']}/file").status_code, 404)
        self.assertEqual(self.client.get('/signed-fiches/export.zip?year=2026&month=9').status_code, 404)
        self.assertEqual(self.delete(row['id']).status_code, 404)
        audit = self.rows('SELECT * FROM permanent_deletion_events')
        self.assertEqual(len(audit), 1)
        self.assertEqual((audit[0]['resource_type'], audit[0]['actor_username']), ('signed_fiche', 'super'))
        self.assertNotIn('Synthetic Employee', audit[0]['metadata'])
        self.assertEqual(self.rows('SELECT event FROM signed_fiche_events ORDER BY id')[-1]['event'], 'permanently_deleted')

    def test_deletion_denies_other_roles_and_stale_super_admin_sessions(self):
        row = self.ready()
        for name in ('admin', 'manager', 'employee', 'other'):
            self.actor = self.identity(name)
            self.assertFalse(self.client.get('/signed-fiches?year=2026').json()['can_delete'])
            self.assertEqual(self.delete(row['id']).status_code, 403)
        self.actor = None
        self.assertEqual(self.delete(row['id']).status_code, 401)
        self.actor = self.identity('super')
        with self.service.store.transaction() as cursor:
            cursor.execute("UPDATE users SET role_tags=%s WHERE username='super'", ('["Admin"]',))
        self.assertEqual(self.delete(row['id']).status_code, 403)
        with self.service.store.transaction() as cursor:
            cursor.execute("UPDATE users SET role_tags=%s,session_version=1 WHERE username='super'", ('["SuperAdmin"]',))
        self.assertEqual(self.delete(row['id']).status_code, 401)
        self.assertTrue(self.service.path(row).exists())
        self.assertEqual(len(self.rows()), 1)
        self.assertEqual(self.rows('SELECT * FROM permanent_deletion_events'), [])

    def test_deletion_rejects_unfinished_requests_and_wrong_confirmation(self):
        row = self.ready()
        self.actor = self.identity('super')
        self.assertEqual(self.delete(row['id'], 'wrong').status_code, 400)
        self.assertEqual(self.client.request('DELETE', f"/signed-fiches/{row['id']}", json={}).status_code, 422)
        for status in ('pending', 'sent', 'partially_signed', 'completed', 'uncertain', 'cancelled', 'expired', 'declined'):
            with self.service.store.transaction() as cursor:
                cursor.execute('UPDATE signed_fiche_documents SET status=%s WHERE id=%s', (status, row['id']))
            self.assertEqual(self.delete(row['id']).status_code, 409)
        self.assertTrue(self.service.path(row).exists())
        self.assertEqual(self.rows('SELECT * FROM permanent_deletion_events'), [])

    def test_deletion_audit_failure_rolls_back_before_file_removal(self):
        row = self.ready()
        self.actor = self.identity('super')
        with patch('backend.signed_fiches.append_deletion_event', side_effect=RuntimeError('Synthetic audit failure')):
            self.assertEqual(self.delete(row['id']).status_code, 500)
        self.assertEqual(self.service.read(self.rows()[0]), self.pdf)
        self.assertEqual(self.rows('SELECT * FROM permanent_deletion_events'), [])
        self.fail_audit = True
        self.assertEqual(self.delete(row['id']).status_code, 500)
        self.assertEqual(self.rows('SELECT * FROM permanent_deletion_events'), [])
        self.assertTrue(self.service.path(row).exists())

    def test_file_cleanup_failure_is_explicit_and_removed_row_cannot_reappear(self):
        request = self.prepare()
        self.assertEqual(self.notify(request).status_code, 204)
        claimed = self.service.store.claim()
        self.service.save(claimed, self.pdf)
        self.actor = self.identity('super')
        with patch('backend.signed_fiches.Path.unlink', side_effect=OSError('Synthetic permission failure')):
            result = self.delete(request['id'])
        self.assertEqual(result.status_code, 200, result.text)
        self.assertTrue(result.json()['cleanup_pending'])
        self.assertEqual(self.rows(), [])
        self.assertEqual(self.notify(request).status_code, 404)
        self.assertIsNone(self.service.store.claim())

    def test_deletion_rejects_archive_path_escape_and_keeps_external_file(self):
        row = self.ready()
        self.actor = self.identity('super')
        with patch.object(self.service, 'path', side_effect=ValueError('Synthetic invalid path')):
            self.assertEqual(self.delete(row['id']).status_code, 409)
        self.assertTrue(self.service.path(row).exists())
        self.assertEqual(len(self.rows()), 1)

    def test_successor_cannot_inherit_old_signature_even_if_they_sent_it(self):
        row = self.ready(manager='former', sender='manager')
        self.actor = self.identity('manager')
        for query in ('', '&username=employee'):
            data = self.client.get('/signed-fiches?year=2026&month=9' + query).json()
            self.assertEqual(data['total'], 0)
            self.assertEqual(data['items'], [])
        self.assertEqual(self.client.get(f"/signed-fiches/{row['id']}/file").status_code, 403)
        for actor in ('employee', 'admin', 'super'):
            self.actor = self.identity(actor)
            self.assertEqual(self.client.get('/signed-fiches?year=2026&month=9').json()['total'], 1)
            self.assertEqual(self.client.get(f"/signed-fiches/{row['id']}/file").status_code, 200)

    def test_unknown_historical_signer_is_not_inferred_from_sender(self):
        row = self.ready(manager=None, sender='manager')
        self.actor = self.identity('manager')
        self.assertEqual(self.client.get('/signed-fiches?year=2026&month=9').json()['total'], 0)
        self.assertEqual(self.client.get(f"/signed-fiches/{row['id']}/file").status_code, 403)
        self.actor = self.identity('employee')
        self.assertEqual(self.client.get(f"/signed-fiches/{row['id']}/file").status_code, 200)

    def test_original_signer_loses_history_when_employee_changes_manager(self):
        row = self.ready()
        with self.service.store.transaction() as cursor:
            cursor.execute("UPDATE users SET manager_username='former' WHERE username='employee'")
        for actor in ('manager', 'former'):
            self.actor = self.identity(actor)
            self.assertEqual(self.client.get('/signed-fiches?year=2026&month=9').json()['total'], 0)
            self.assertEqual(self.client.get(f"/signed-fiches/{row['id']}/file").status_code, 403)

    def test_signer_account_rename_preserves_attribution_without_rewriting_pdf(self):
        row = self.ready()
        with self.service.store.transaction() as cursor:
            rename_signed_fiche_owner(cursor, 'manager', 'renamed')
            cursor.execute("UPDATE users SET username='renamed' WHERE username='manager'")
            cursor.execute("UPDATE users SET manager_username='renamed' WHERE username='employee'")
        self.actor = self.identity('renamed')
        self.assertEqual(self.rows()[0]['manager_username'], 'renamed')
        self.assertEqual(self.client.get('/signed-fiches?year=2026&month=9').json()['total'], 1)
        self.assertEqual(self.client.get(f"/signed-fiches/{row['id']}/file").content, self.pdf)

    def test_callback_rejects_forged_identity_and_document(self):
        request = self.prepare(); self.service.store.accepted(request['id'], 'fixture_doc')
        for changes in [dict(reference_id='forged'),dict(namespace='wrong'),dict(provider='wrong'),dict(document_id='different')]:
            self.assertIn(self.notify(request, **changes).status_code, {404,409})
        self.assertEqual(self.rows()[0]['status'], 'sent')

    def test_first_signature_does_not_archive_and_boolean_must_be_true(self):
        request = self.prepare()
        self.assertEqual(self.notify(request,event='signer.signed',all_signed=False).status_code, 204)
        self.assertEqual(self.rows()[0]['status'], 'partially_signed')
        self.assertIsNone(self.service.store.claim())
        self.assertEqual(self.notify(request,all_signed='true').status_code, 400)

    def test_completion_before_acceptance_and_duplicates_never_regress(self):
        request = self.prepare()
        self.assertEqual(self.notify(request).status_code, 204)
        self.service.store.accepted(request['id'], 'fixture_doc')
        count = len(self.rows('SELECT * FROM signed_fiche_events'))
        for _ in range(3): self.assertEqual(self.notify(request).status_code, 204)
        self.notify(request,event='signer.signed',all_signed=False)
        self.notify(request,event='document.expired',all_signed=False)
        self.assertEqual(self.rows()[0]['status'], 'completed')
        self.assertEqual(len(self.rows('SELECT * FROM signed_fiche_events')), count)

    def test_failed_audit_rolls_back_callback_state(self):
        request = self.prepare(); self.fail_audit = True
        self.assertEqual(self.notify(request).status_code, 500)
        self.assertEqual(self.rows()[0]['status'], 'pending')

    def test_claim_lease_retry_and_stale_worker_are_serialized(self):
        request = self.prepare(); self.notify(request)
        first = self.service.store.claim()
        self.assertIsNone(self.service.store.claim())
        self.now += timedelta(seconds=181)
        second = self.service.store.claim()
        self.service.store.downloaded(first, '0'*64, 123)
        self.assertEqual(self.rows()[0]['status'], 'completed')
        self.service.store.failed(second)
        row = self.rows()[0]
        self.assertIsNone(row['lease_token']); self.assertTrue(row['last_error'])
        self.assertIsNone(self.service.store.claim())
        self.now += timedelta(hours=1)
        self.assertIsNotNone(self.service.store.claim())

    def test_original_pdf_preserved_private_and_month_is_attendance_month(self):
        row = self.ready()
        self.assertEqual(self.service.path(row).relative_to(self.service.root).parts[:2], ('2026','09'))
        response = self.client.get(f"/signed-fiches/{row['id']}/file")
        self.assertEqual(response.content, self.pdf)
        self.assertEqual(hashlib.sha256(response.content).hexdigest(), row['sha256'])
        self.assertIn('no-store', response.headers['cache-control'])
        self.assertNotIn('signed_pdf_url', response.text if response.headers.get('content-type')=='application/json' else '')

    def test_account_roles_and_fresh_session_are_checked_on_download(self):
        row = self.ready(); path = f"/signed-fiches/{row['id']}/file"
        for name in ['employee','manager','admin','super']:
            self.actor = self.identity(name); self.assertEqual(self.client.get(path).status_code, 200)
        for name in ['other','former']:
            self.actor = self.identity(name); self.assertEqual(self.client.get(path).status_code, 403)
        self.actor = self.identity('admin')
        self.rows("UPDATE users SET session_version=1 WHERE username='admin'")
        self.assertEqual(self.client.get(path).status_code, 401)

    def test_current_manager_only_and_account_rename_preserve_owner(self):
        row = self.ready()
        self.rows("UPDATE users SET manager_username='former' WHERE username='employee'")
        self.actor=self.identity('manager')
        self.assertEqual(self.client.get(f"/signed-fiches/{row['id']}/file").status_code,403)
        with self.service.store.transaction() as cursor: rename_signed_fiche_owner(cursor,'employee','new-employee')
        self.assertEqual(self.rows()[0]['employee_username'],'new-employee')
        self.assertEqual(self.service.read(self.rows()[0]), self.pdf)

    def test_zip_is_admin_only_separates_specimens_and_rejects_partial_archive(self):
        self.ready()
        self.ready(mode='sandbox-only')
        path='/signed-fiches/export.zip?year=2026&month=9'
        self.actor=self.identity('employee'); self.assertEqual(self.client.get(path).status_code,403)
        self.actor=self.identity('admin'); response=self.client.get(path)
        self.assertEqual(response.status_code,200)
        with ZipFile(io.BytesIO(response.content)) as archive:
            self.assertEqual(len(archive.namelist()),1)
            self.assertTrue(archive.namelist()[0].startswith('2026/09/'))
            self.assertEqual(archive.read(archive.namelist()[0]),self.pdf)
        self.notify(self.prepare())
        self.assertEqual(self.client.get(path).status_code,409)

    def test_corrupted_file_never_downloads_or_exports(self):
        row=self.ready(); self.service.path(row).write_bytes(b'corrupted')
        self.assertEqual(self.client.get(f"/signed-fiches/{row['id']}/file").status_code,503)
        self.assertEqual(self.client.get('/signed-fiches/export.zip?year=2026&month=9').status_code,503)

    def test_yearly_zip_groups_all_pages_by_month_and_separates_year_and_mode(self):
        expected = [self.ready(month=1, name='Synthetic ../ Person'), self.ready(month=12)]
        for index in range(20):
            expected.append(self.ready(month=9, name=f'Synthetic {index}'))
        self.ready(year=2025)
        self.ready(mode='sandbox-only')
        response = self.client.get('/signed-fiches/export.zip?year=2026')
        self.assertEqual(response.status_code, 200)
        self.assertIn('ILCI_RH_2026.zip', response.headers['content-disposition'])
        self.assertIn('no-store', response.headers['cache-control'])
        with ZipFile(io.BytesIO(response.content)) as archive:
            names = archive.namelist()
            self.assertEqual(len(names), 22)
            self.assertEqual({name.split('/')[1] for name in names}, {'01', '09', '12'})
            self.assertTrue(all(name.startswith('ILCI_RH_2026/') and '..' not in name.split('/') for name in names))
            for name in names:
                self.assertEqual(archive.read(name), self.pdf)
        audited = self.rows("SELECT archive_id FROM signed_fiche_events WHERE event='yearly_export'")
        self.assertEqual({r['archive_id'] for r in audited}, {r['id'] for r in expected})
        specimen = self.client.get('/signed-fiches/export.zip?year=2026&mode=sandbox-only')
        self.assertEqual(specimen.status_code, 200)
        self.assertIn('SPECIMEN_ILCI_RH_2026.zip', specimen.headers['content-disposition'])
        with ZipFile(io.BytesIO(specimen.content)) as archive:
            self.assertEqual(len(archive.namelist()), 1)
            self.assertTrue(archive.namelist()[0].startswith('SPECIMEN_ILCI_RH_2026/09/'))

    def test_yearly_export_rejects_non_admins_and_stale_admin_role(self):
        self.ready()
        path = '/signed-fiches/export.zip?year=2026'
        for username in ('employee', 'manager', 'other'):
            self.actor = self.identity(username)
            self.assertEqual(self.client.get(path).status_code, 403)
        self.actor = self.identity('super')
        self.assertEqual(self.client.get(path).status_code, 200)
        self.actor = self.identity('admin')
        with closing(connect_fixture(self.path)) as db, db:
            db.execute('UPDATE users SET role_tags=? WHERE username=?', ('["Employe"]', 'admin'))
        self.assertEqual(self.client.get(path).status_code, 403)

    def test_yearly_export_never_returns_partial_zip_or_keeps_partial_audit(self):
        row = self.ready(month=1)
        late = self.prepare(month=12)
        self.assertEqual(self.notify(late).status_code, 204)
        path = '/signed-fiches/export.zip?year=2026'
        self.assertEqual(self.client.get(path).status_code, 409)
        self.assertEqual(self.rows("SELECT * FROM signed_fiche_events WHERE event='yearly_export'"), [])
        with self.service.store.transaction() as cursor:
            cursor.execute("UPDATE signed_fiche_documents SET status='cancelled' WHERE id=%s", (late['id'],))
        self.service.path(row).write_bytes(b'synthetic corrupted PDF')
        self.assertEqual(self.client.get(path).status_code, 503)
        self.assertEqual(self.rows("SELECT * FROM signed_fiche_events WHERE event='yearly_export'"), [])

    def test_yearly_export_limits_and_audit_failure_return_no_zip(self):
        row = self.ready()
        path = '/signed-fiches/export.zip?year=2026'
        with patch('backend.signed_fiches.MAX_ARCHIVE_BYTES', 1):
            self.assertEqual(self.client.get(path).status_code, 413)
        self.fail_audit = True
        self.assertEqual(self.client.get(path).status_code, 500)
        self.fail_audit = False
        self.assertEqual(self.rows("SELECT * FROM signed_fiche_events WHERE event='yearly_export'"), [])
        self.assertEqual(self.service.read(row), self.pdf)

    def test_atomic_archive_refuses_different_replacement(self):
        row=self.ready()
        with self.assertRaises(ValueError): self.service.save(row,self.pdf+b'changed')
        self.assertEqual(self.service.read(row),self.pdf)

    def test_list_filters_and_twenty_row_pagination(self):
        for i in range(21): self.prepare(name=f'Synthetic {i}')
        self.prepare(employee='other',mode='sandbox-only')
        first=self.client.get('/signed-fiches?year=2026&month=9').json()
        second=self.client.get('/signed-fiches?year=2026&month=9&page=2').json()
        self.assertEqual(first['total'],21); self.assertEqual(len(first['items']),20); self.assertEqual(len(second['items']),1)
        self.actor=self.identity('other')
        self.assertEqual(self.client.get('/signed-fiches?year=2026&month=9').json()['total'],0)
        self.assertEqual(self.client.get('/signed-fiches?year=2026&month=9&username=employee').status_code,403)

    def test_annual_listing_groups_months_and_keeps_manager_scope(self):
        self.prepare(month=2)
        self.prepare(month=10)
        self.prepare(month=10, manager='former')
        self.prepare(month=12, year=2025)
        self.prepare(month=11, mode='sandbox-only')
        annual = self.client.get('/signed-fiches?year=2026').json()
        self.assertEqual(annual['total'], 3)
        self.assertEqual([row['fiche_month'] for row in annual['items']], [10, 10, 2])
        self.assertEqual(self.client.get('/signed-fiches?year=2026&month=2').json()['total'], 1)
        self.actor = self.identity('manager')
        self.assertEqual(self.client.get('/signed-fiches?year=2026').json()['total'], 2)
        self.actor = self.identity('other')
        self.assertEqual(self.client.get('/signed-fiches?year=2026').json()['total'], 0)
        self.assertEqual(self.client.get('/signed-fiches?year=2026&month=13').status_code, 422)

    def test_body_limit_and_malformed_callback(self):
        self.assertEqual(self.client.post('/esign/callback',content=b'x'*32769).status_code,413)
        for body in ['[]','null','invalid']:
            self.assertEqual(self.client.post('/esign/callback',content=body).status_code,400)

    def test_url_validation_rejects_ssrf_preview_and_credentials(self):
        for url in ['http://esign.clawshow.ai/esign/fixture_doc/signed.pdf','https://evil.test/esign/fixture_doc/signed.pdf',
                    'https://esign.clawshow.ai/esign/fixture_doc/preview.pdf','https://u:p@esign.clawshow.ai/esign/fixture_doc/signed.pdf',
                    'https://esign.clawshow.ai:444/esign/fixture_doc/signed.pdf','https://127.0.0.1/esign/fixture_doc/signed.pdf']:
            with self.assertRaises(HTTPException): validate_signed_url(url,'fixture_doc')
        self.assertTrue(validate_signed_url('https://esign.clawshow.ai/esign/fixture_doc/signed.pdf','fixture_doc'))

    def test_download_validates_pdf_and_does_not_rewrite_bytes(self):
        response=Mock(); response.status=200; response.headers={}; response.read.side_effect=[self.pdf,b'']
        response.__enter__=Mock(return_value=response); response.__exit__=Mock(return_value=False)
        with patch('backend.signed_fiches.urllib.request.build_opener') as opener:
            opener.return_value.open.return_value=response
            self.assertEqual(fetch_signed_pdf({'signed_pdf_url':'https://esign.clawshow.ai/esign/fixture_doc/signed.pdf','document_id':'fixture_doc'}),self.pdf)

    def test_unconfigured_callback_does_not_create_tracking_row(self):
        with patch.dict(os.environ,{'PRESENCE_ESIGN_CALLBACK_URL':''}):
            self.assertIsNone(self.prepare()); self.assertEqual(self.rows(),[])

    def test_retry_is_admin_only_and_never_sends_signature(self):
        request=self.prepare(); self.notify(request)
        row=self.service.store.claim(); self.service.store.failed(row)
        self.actor=self.identity('employee')
        self.assertEqual(self.client.post(f"/signed-fiches/{request['id']}/retry").status_code,403)
        self.actor=self.identity('admin')
        self.assertEqual(self.client.post(f"/signed-fiches/{request['id']}/retry").status_code,200)
        self.assertIsNotNone(self.service.store.claim())

    def test_download_rejects_oversize_and_non_pdf_response(self):
        for headers, chunks in [({'Content-Length':str(26*1024*1024)},[]),({},[b'<html>error</html>',b''])]:
            response=Mock(); response.status=200; response.headers=headers; response.read.side_effect=chunks
            response.__enter__=Mock(return_value=response); response.__exit__=Mock(return_value=False)
            with patch('backend.signed_fiches.urllib.request.build_opener') as opener:
                opener.return_value.open.return_value=response
                with self.assertRaises(ValueError): fetch_signed_pdf({'signed_pdf_url':'https://esign.clawshow.ai/esign/fixture_doc/signed.pdf','document_id':'fixture_doc'})


if __name__ == '__main__': unittest.main()
