"""Isolated SQL and private files for administrator-managed invoice types."""
import base64
import hashlib
from pathlib import Path
import sqlite3
import tempfile
import unittest
from unittest.mock import patch

import pymysql
from fastapi import FastAPI, HTTPException
from fastapi.testclient import TestClient
from backend.documents import register_document_routes
from backend.document_categories import label_key


class Cursor:
    def __init__(self, fixture): self.fixture = fixture; self.raw = fixture.sql.cursor()
    def __enter__(self): return self
    def __exit__(self, *_): self.raw.close()
    @property
    def lastrowid(self): return self.raw.lastrowid
    @property
    def rowcount(self): return self.raw.rowcount
    def execute(self, query, params=()):
        if 'INSERT INTO document_category_events' in query and self.fixture.fail_audit:
            raise RuntimeError('Synthetic audit failure')
        query = query.replace(' FOR UPDATE', '').replace('%s', '?').replace('FROM_UNIXTIME(?)', "datetime(?, 'unixepoch')")
        try: return self.raw.execute(query, tuple(params))
        except sqlite3.IntegrityError as error: raise pymysql.err.IntegrityError(1062, 'Fixture duplicate') from error
    def fetchone(self):
        row = self.raw.fetchone()
        return dict(row) if row else None
    def fetchall(self): return [dict(row) for row in self.raw.fetchall()]


class Connection:
    def __init__(self, fixture): self.fixture = fixture
    def cursor(self): return Cursor(self.fixture)
    def commit(self): self.fixture.sql.commit()
    def rollback(self): self.fixture.sql.rollback()
    def close(self): pass


class DocumentCategoryTests(unittest.TestCase):
    def setUp(self):
        self.sql = sqlite3.connect(':memory:', check_same_thread=False)
        self.sql.row_factory = sqlite3.Row
        self.addCleanup(self.sql.close)
        self.temp = tempfile.TemporaryDirectory(); self.addCleanup(self.temp.cleanup)
        self.storage = Path(self.temp.name)
        self.fail_audit = False
        self.sql.executescript('''
            CREATE TABLE users (username TEXT PRIMARY KEY);
            INSERT INTO users VALUES ('admin'),('employee'),('manager'),('super');
            CREATE TABLE document_categories (code TEXT PRIMARY KEY,label TEXT,label_key TEXT UNIQUE,
                is_active INT DEFAULT 1,is_system INT DEFAULT 0,version INT DEFAULT 1);
            CREATE TABLE document_category_events (id INTEGER PRIMARY KEY,category_code TEXT,receipt_id INT,
                actor_username TEXT,action TEXT,before_json TEXT,after_json TEXT);
            CREATE TABLE transport_receipts (id INTEGER PRIMARY KEY,employee_username TEXT,uploaded_by TEXT,
                original_filename TEXT,stored_filename TEXT,mime_type TEXT,size_bytes INT,document_type TEXT,other_type TEXT,
                receipt_month INT,receipt_year INT,document_number TEXT,created_at TEXT,archived INT DEFAULT 0,
                archived_at TEXT,archived_by TEXT);
        ''')
        self.sql.executemany('INSERT INTO document_categories (code,label,label_key,is_system) VALUES (?,?,?,1)',
            [(code, label, label_key(label)) for code, label in [('transport', 'Transport'), ('autre', 'Autre')]])
        self.sql.commit()
        self.accounts = {name: {'username': name, 'role_tags': tags, 'session_version': 0, 'is_active': True} for name, tags in (
            ('admin', ['Admin']), ('employee', ['Employe']), ('manager', ['Responsable']), ('super', ['SuperAdmin']))}
        self.actor = self.accounts['admin'].copy()
        app = FastAPI()
        def authenticated():
            if self.actor is None: raise HTTPException(401, 'Authentication required')
            return self.actor
        register_document_routes(app, get_db_connection=lambda: Connection(self), get_current_user=authenticated,
            require_roles=lambda *_: authenticated, load_user=lambda _c, name: self.accounts.get(name),
            user_can_manage=lambda *_: False, all_users=lambda _: list(self.accounts.values()),
            storage_dir=self.storage, max_file_bytes=1024 * 1024)
        self.client = TestClient(app, raise_server_exceptions=False); self.addCleanup(self.client.close)
        self.addCleanup(patch.stopall)
        patch('backend.documents.load_display_names', return_value={'employee': 'Synthetic Employee', 'admin': 'Synthetic Admin'}).start()
        patch('backend.documents.allocate_document_number', return_value='20261028-0001').start()
        self.pdf = b'%PDF-1.4\nSynthetic category test\n%%EOF'

    def create(self, label='Hôtel'):
        return self.client.post('/document-categories', json={'label': label})
    def update(self, category, **values):
        return self.client.patch('/document-categories/' + category['code'], json={'expected_version': category['version'], **values})
    def upload(self, code='transport', other=None):
        return self.client.post('/transport-receipts', json={'employee_username': 'employee', 'document_type': code,
            'other_type': other, 'receipt_month': 10, 'receipt_year': 2026,
            'original_filename': 'invoice.pdf', 'file_b64': base64.b64encode(self.pdf).decode()})
    def classify(self, receipt, code, other=None):
        return self.client.patch(f"/transport-receipts/{receipt['id']}/category", json={'document_type': code,
            'other_type': other, 'expected_document_type': receipt['document_type'], 'expected_other_type': receipt['other_type']})
    def total(self, table): return self.sql.execute(f'SELECT COUNT(*) FROM {table}').fetchone()[0]

    def test_admin_creates_global_type_and_employee_can_select_and_upload(self):
        response = self.create(); self.assertEqual(response.status_code, 201, response.text)
        category = response.json(); self.assertEqual(len(category['code']), 20)
        self.actor = self.accounts['employee']
        listing = self.client.get('/document-categories')
        self.assertEqual(listing.status_code, 200)
        self.assertFalse(listing.json()['can_manage'])
        self.assertIn(category, listing.json()['categories'])
        response = self.upload(category['code'])
        self.assertEqual(response.status_code, 200, response.text)
        self.assertEqual(response.json()['receipt']['category_label'], 'Hôtel')
        self.assertEqual(next(self.storage.iterdir()).read_bytes(), self.pdf)

    def test_super_admin_inherits_category_management(self):
        self.actor = self.accounts['super']
        self.assertEqual(self.create().status_code, 201)

    def test_employee_and_responsable_cannot_manage_or_reclassify(self):
        category = self.create().json()
        receipt = self.upload().json()['receipt']
        for name in ('employee', 'manager'):
            self.actor = self.accounts[name]
            self.assertEqual(self.create('Restauration').status_code, 403)
            self.assertEqual(self.update(category, is_active=False).status_code, 403)
            self.assertEqual(self.classify(receipt, category['code']).status_code, 403)
        self.assertEqual(self.total('document_category_events'), 1)

    def test_live_role_and_session_revocation_are_enforced(self):
        for changes, expected in (({'role_tags': ['Employe']}, 403), ({'is_active': False}, 401),
                                  ({'session_version': 1}, 401), ({'login_deleted_at': '2026-10-28'}, 401)):
            self.accounts['admin'] = {**self.actor, **changes}
            self.assertEqual(self.create().status_code, expected)
        self.assertEqual(self.total('document_categories'), 2)

    def test_labels_are_trimmed_unique_across_case_accents_and_disabled_types(self):
        category = self.create('  Hôtel   Paris  ').json()
        self.assertEqual(category['label'], 'Hôtel Paris')
        self.assertEqual(self.create('HOTEL PARIS').status_code, 409)
        self.assertEqual(self.create('transport').status_code, 409)
        self.assertEqual(self.update(category, is_active=False).status_code, 200)
        self.assertEqual(self.create('hotel paris').status_code, 409)

    def test_blank_hidden_overlong_and_reserved_labels_are_rejected(self):
        for label in ('', '   ', 'a' * 81, 'a\u200bb', 'Tous', 'Non renseigné'):
            self.assertEqual(self.create(label).status_code, 422)
        self.assertEqual(self.total('document_categories'), 2)

    def test_disable_keeps_old_files_labels_and_filters_but_rejects_uploads(self):
        category = self.create().json()
        receipt = self.upload(category['code']).json()['receipt']
        self.assertEqual(self.update(category, is_active=False).status_code, 200)
        self.actor = self.accounts['employee']
        response = self.upload(category['code'])
        self.assertEqual(response.status_code, 409)
        self.assertEqual(self.total('transport_receipts'), 1)
        self.assertEqual(self.client.get(f"/transport-receipts/{receipt['id']}/file").content, self.pdf)
        self.assertEqual(self.client.get('/transport-receipts').json()['receipts'][0]['category_label'], 'Hôtel')

    def test_rename_preserves_code_and_receipt_bytes_and_reactivation(self):
        category = self.create().json(); receipt = self.upload(category['code']).json()['receipt']
        updated = self.update(category, label='Hébergement').json()
        self.assertEqual(updated['code'], category['code'])
        self.assertEqual(self.client.get('/transport-receipts').json()['receipts'][0]['category_label'], 'Hébergement')
        disabled = self.update(updated, is_active=False).json()
        self.assertEqual(self.update(disabled, is_active=True).status_code, 200)
        self.assertEqual(self.client.get(f"/transport-receipts/{receipt['id']}/file").content, self.pdf)

    def test_stale_category_and_receipt_changes_do_not_overwrite(self):
        category = self.create().json()
        self.assertEqual(self.update(category, label='Repas').status_code, 200)
        self.assertEqual(self.update(category, is_active=False).status_code, 409)
        receipt = self.upload().json()['receipt']
        self.assertEqual(self.classify(receipt, category['code']).status_code, 200)
        self.assertEqual(self.classify(receipt, 'autre', 'Changed').status_code, 409)

    def test_classification_changes_only_metadata_and_is_audited(self):
        category = self.create().json(); receipt = self.upload().json()['receipt']
        before = dict(self.sql.execute('SELECT * FROM transport_receipts').fetchone())
        file = next(self.storage.iterdir()); digest = hashlib.sha256(file.read_bytes()).hexdigest()
        response = self.classify(receipt, category['code'])
        self.assertEqual(response.status_code, 200, response.text)
        after = dict(self.sql.execute('SELECT * FROM transport_receipts').fetchone())
        self.assertEqual({k: v for k, v in before.items() if k not in {'document_type', 'other_type'}},
                         {k: v for k, v in after.items() if k not in {'document_type', 'other_type'}})
        self.assertEqual(hashlib.sha256(file.read_bytes()).hexdigest(), digest)
        self.assertEqual(self.sql.execute("SELECT action FROM document_category_events ORDER BY id DESC LIMIT 1").fetchone()[0], 'reclassified')

    def test_audit_failure_rolls_back_creation_and_reclassification(self):
        category = self.create().json(); receipt = self.upload().json()['receipt']
        self.fail_audit = True
        self.assertEqual(self.create('Repas').status_code, 500)
        self.assertEqual(self.total('document_categories'), 3)
        self.assertEqual(self.classify(receipt, category['code']).status_code, 500)
        self.assertEqual(self.sql.execute('SELECT document_type FROM transport_receipts').fetchone()[0], 'transport')

    def test_builtin_and_legacy_other_and_unclassified_stay_compatible(self):
        self.assertEqual(self.update({'code': 'transport', 'version': 1}, label='New').status_code, 409)
        self.assertEqual(self.upload('autre').status_code, 400)
        response = self.upload('autre', 'Taxi exceptionnel')
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.json()['receipt']['category_label'], 'Autre · Taxi exceptionnel')
        self.sql.execute('UPDATE transport_receipts SET document_type=NULL,other_type=NULL'); self.sql.commit()
        self.assertEqual(self.client.get('/transport-receipts').json()['receipts'][0]['category_label'], 'Non renseigné')

    def test_unknown_type_never_writes_files_or_receipts(self):
        self.assertEqual(self.upload('c_unknown').status_code, 400)
        self.assertEqual(self.total('transport_receipts'), 0)
        self.assertEqual(list(self.storage.iterdir()), [])


if __name__ == '__main__': unittest.main()
