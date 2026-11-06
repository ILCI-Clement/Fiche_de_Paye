"""Atomic file batches against isolated synthetic SQL and private temporary files."""
import unittest
from unittest.mock import patch

import test_document_categories as documents
import test_signed_fiches as signed


class SignedBatchTests(unittest.TestCase):
    setUp = signed.SignedFicheTests.setUp
    authenticated = signed.SignedFicheTests.authenticated
    load_user = signed.SignedFicheTests.load_user
    identity = signed.SignedFicheTests.identity
    rows = signed.SignedFicheTests.rows
    prepare = signed.SignedFicheTests.prepare
    payload = signed.SignedFicheTests.payload
    notify = signed.SignedFicheTests.notify
    ready = signed.SignedFicheTests.ready

    def batch(self, ids, action='archive', archived=False, confirmation=None):
        return self.client.post('/signed-fiches/batch', json=dict(ids=ids, action=action,
            expected_archived=archived, confirmation=confirmation))

    def test_archive_restore_preserves_pdf_and_history(self):
        rows = [self.ready(), self.ready(month=10)]
        ids = [row['id'] for row in rows]
        response = self.batch(ids + ids)
        self.assertEqual(response.status_code, 200, response.text)
        self.assertEqual(response.json()['count'], 2)
        self.assertEqual(self.client.get('/signed-fiches?year=2026').json()['total'], 0)
        self.assertEqual(self.client.get('/signed-fiches?year=2026&archived=true').json()['total'], 2)
        for row in rows:
            self.assertEqual(self.service.read(row), self.pdf)
            self.assertEqual(self.client.get(f"/signed-fiches/{row['id']}/file").status_code, 200)
        # Manual folder state must not remove files from the complete annual export.
        self.assertEqual(self.client.get('/signed-fiches/export.zip?year=2026').status_code, 200)
        self.assertEqual(self.batch(ids, 'restore', True).status_code, 200)
        self.assertEqual(self.client.get('/signed-fiches?year=2026').json()['total'], 2)
        events = self.rows("SELECT event FROM signed_fiche_events WHERE event IN ('user_archived','user_restored')")
        self.assertEqual(len(events), 4)

    def test_invalid_member_or_stale_state_changes_nothing(self):
        row = self.ready()
        pending = self.prepare()
        for ids, expected in (([row['id'], pending['id']], 409), ([row['id'], 'f' * 32], 404)):
            result = self.batch(ids)
            self.assertEqual(result.status_code, expected, result.text)
            self.assertFalse(self.rows('SELECT is_archived FROM signed_fiche_documents WHERE id=%s', (row['id'],))[0]['is_archived'])
        self.assertEqual(self.batch([row['id']], archived=True).status_code, 409)
        self.assertEqual(self.rows("SELECT * FROM signed_fiche_events WHERE event='user_archived'"), [])

    def test_permissions_are_rechecked_for_all_actions(self):
        row = self.ready()
        for name in ('employee', 'manager', 'other'):
            self.actor = self.identity(name)
            self.assertEqual(self.batch([row['id']]).status_code, 403)
        self.actor = self.identity('admin')
        self.assertEqual(self.batch([row['id']], 'delete', confirmation='SUPPRIMER').status_code, 403)
        self.actor = self.identity('super')
        with self.service.store.transaction() as cursor:
            cursor.execute("UPDATE users SET role_tags=%s WHERE username='super'", ('["Employe"]',))
        self.assertEqual(self.batch([row['id']]).status_code, 403)
        self.assertTrue(self.service.path(row).exists())

    def test_short_confirmation_deletes_only_selected_and_audits(self):
        rows = [self.ready(), self.ready()]
        self.actor = self.identity('super')
        self.assertEqual(self.batch([rows[0]['id']], 'delete', confirmation='wrong').status_code, 400)
        response = self.batch([rows[0]['id']], 'delete', confirmation='SUPPRIMER')
        self.assertEqual(response.status_code, 200, response.text)
        self.assertFalse(self.service.path(rows[0]).exists())
        self.assertTrue(self.service.path(rows[1]).exists())
        self.assertEqual(len(self.rows('SELECT * FROM permanent_deletion_events')), 1)
        response = self.client.request('DELETE', f"/signed-fiches/{rows[1]['id']}", json={'confirmation': 'SUPPRIMER'})
        self.assertEqual(response.status_code, 200, response.text)

    def test_audit_failure_rolls_back_entire_batch_before_removing_files(self):
        rows = [self.ready(), self.ready()]
        self.fail_audit = True
        self.assertEqual(self.batch([r['id'] for r in rows]).status_code, 500)
        self.assertTrue(all(not r['is_archived'] for r in self.rows()))
        self.actor = self.identity('super')
        self.assertEqual(self.batch([r['id'] for r in rows], 'delete', confirmation='SUPPRIMER').status_code, 500)
        self.assertEqual(len(self.rows()), 2)
        self.assertEqual(self.rows('SELECT * FROM permanent_deletion_events'), [])
        self.assertTrue(all(self.service.path(r).exists() for r in rows))

    def test_cleanup_failure_is_reported_after_commit(self):
        row = self.ready()
        self.actor = self.identity('super')
        with patch('pathlib.Path.unlink', side_effect=OSError('Synthetic failure')):
            response = self.batch([row['id']], 'delete', confirmation='SUPPRIMER')
        self.assertEqual(response.status_code, 200, response.text)
        self.assertTrue(response.json()['cleanup_pending'])
        self.assertEqual(self.rows(), [])
        self.assertEqual(self.client.get(f"/signed-fiches/{row['id']}/file").status_code, 404)

    def test_payload_limits_and_path_validation(self):
        for ids, expected in (([], 422), (['f' * 32] * 101, 422), (['invalid'], 400)):
            self.assertEqual(self.batch(ids).status_code, expected)
        row = self.ready()
        self.actor = self.identity('super')
        with patch.object(self.service, 'path', side_effect=ValueError('Invalid path')):
            self.assertEqual(self.batch([row['id']], 'delete', confirmation='SUPPRIMER').status_code, 409)
        self.assertEqual(len(self.rows()), 1)


class DocumentBatchTests(unittest.TestCase):
    upload = documents.DocumentCategoryTests.upload
    total = documents.DocumentCategoryTests.total

    def setUp(self):
        documents.DocumentCategoryTests.setUp(self)
        self.sql.executescript('''
            CREATE TABLE resource_archive_events(id INTEGER PRIMARY KEY,actor_username TEXT,
                resource_type TEXT,resource_id TEXT,is_archived INT);
            CREATE TABLE permanent_deletion_events(id INTEGER PRIMARY KEY,actor_username TEXT,
                resource_type TEXT,resource_id TEXT,metadata TEXT);
        ''')

    def batch(self, ids, action='archive', archived=False, confirmation=None):
        return self.client.post('/transport-receipts/batch', json=dict(ids=ids, action=action,
            expected_archived=archived, confirmation=confirmation))

    def receipt(self):
        response = self.upload()
        self.assertEqual(response.status_code, 200, response.text)
        return response.json()['receipt']

    def test_archive_restore_is_audited_without_changing_file_bytes(self):
        rows = [self.receipt(), self.receipt()]
        ids = [row['id'] for row in rows]
        self.assertEqual(self.batch(ids + ids).json()['count'], 2)
        self.assertEqual(self.sql.execute('SELECT SUM(archived) FROM transport_receipts').fetchone()[0], 2)
        self.assertEqual(self.batch(ids, 'restore', True).status_code, 200)
        self.assertEqual(self.total('resource_archive_events'), 4)
        self.assertTrue(all(path.read_bytes() == self.pdf for path in self.storage.iterdir()))

    def test_own_documents_allowed_but_cross_owner_and_manager_batch_rejected(self):
        rows = [self.receipt(), self.receipt()]
        self.sql.execute("UPDATE transport_receipts SET employee_username='admin',uploaded_by='admin' WHERE id=?", (rows[1]['id'],))
        self.sql.commit()
        self.actor = self.accounts['employee']
        self.assertEqual(self.batch([r['id'] for r in rows]).status_code, 403)
        self.assertEqual(self.total('resource_archive_events'), 0)
        self.assertEqual(self.batch([rows[0]['id']]).status_code, 200)
        self.actor = self.accounts['manager']
        self.assertEqual(self.batch([rows[1]['id']]).status_code, 403)

    def test_missing_or_changed_row_never_partially_updates(self):
        row = self.receipt()
        self.assertEqual(self.batch([row['id'], 999]).status_code, 404)
        self.assertEqual(self.batch([row['id']], archived=True).status_code, 409)
        self.assertEqual(self.total('resource_archive_events'), 0)
        self.assertEqual(self.sql.execute('SELECT archived FROM transport_receipts').fetchone()[0], 0)

    def test_deletion_requires_live_super_admin_archived_files_and_short_confirmation(self):
        rows = [self.receipt(), self.receipt()]
        ids = [r['id'] for r in rows]
        self.assertEqual(self.batch(ids, 'delete', confirmation='SUPPRIMER').status_code, 403)
        self.actor = self.accounts['super'].copy()
        self.assertEqual(self.batch(ids, 'delete', confirmation='SUPPRIMER').status_code, 409)
        self.assertEqual(self.batch(ids).status_code, 200)
        self.assertEqual(self.batch(ids, 'delete', True, 'wrong').status_code, 400)
        self.accounts['super']['role_tags'] = ['Admin']
        self.assertEqual(self.batch(ids, 'delete', True, 'SUPPRIMER').status_code, 403)
        self.accounts['super']['role_tags'] = ['SuperAdmin']
        result = self.batch(ids, 'delete', True, 'SUPPRIMER')
        self.assertEqual(result.status_code, 200, result.text)
        self.assertEqual(self.total('transport_receipts'), 0)
        self.assertEqual(self.total('permanent_deletion_events'), 2)
        self.assertEqual(list(self.storage.iterdir()), [])

    def test_audit_failure_rolls_back_whole_batch_and_retains_files(self):
        rows = [self.receipt(), self.receipt()]
        ids = [r['id'] for r in rows]
        with patch('backend.documents.append_archive_event', side_effect=RuntimeError('Audit unavailable')):
            self.assertEqual(self.batch(ids).status_code, 500)
        self.assertEqual(self.sql.execute('SELECT SUM(archived) FROM transport_receipts').fetchone()[0], 0)
        self.assertEqual(self.batch(ids).status_code, 200)
        self.actor = self.accounts['super']
        with patch('backend.documents.append_deletion_event', side_effect=RuntimeError('Audit unavailable')):
            self.assertEqual(self.batch(ids, 'delete', True, 'SUPPRIMER').status_code, 500)
        self.assertEqual(self.total('transport_receipts'), 2)
        self.assertEqual(len(list(self.storage.iterdir())), 2)

    def test_delete_prevalidates_all_paths_and_checks_limits(self):
        rows = [self.receipt(), self.receipt()]
        ids = [r['id'] for r in rows]
        self.assertEqual(self.batch(ids).status_code, 200)
        self.sql.execute("UPDATE transport_receipts SET stored_filename='../outside.pdf' WHERE id=?", (ids[1],))
        self.sql.commit()
        self.actor = self.accounts['super']
        self.assertEqual(self.batch(ids, 'delete', True, 'SUPPRIMER').status_code, 409)
        self.assertEqual(self.total('transport_receipts'), 2)
        self.assertEqual(self.total('permanent_deletion_events'), 0)
        for values, status in (([], 422), ([1] * 101, 422), ([-1], 400)):
            self.assertEqual(self.batch(values).status_code, status)


if __name__ == '__main__':
    unittest.main()
