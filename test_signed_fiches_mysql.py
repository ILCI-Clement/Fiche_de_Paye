"""Opt-in real SQL integration. Only connection-local TEMPORARY tables are writable."""
from contextlib import closing
import hashlib
import io
import json
import os
from pathlib import Path
import re
import tempfile
import unittest

import pymysql
from fastapi import FastAPI
from fastapi.testclient import TestClient
from pypdf import PdfWriter
from backend.signed_fiche_store import SignedFicheStore, ensure_signed_fiche_schema
from backend.signed_fiches import SignedFiches, validate_signed_url, register_signed_fiche_routes


class TemporaryCursor:
    def __init__(self, db): self.raw=db.cursor()
    def __enter__(self): return self
    def __exit__(self,*_): self.raw.close()
    @property
    def rowcount(self): return self.raw.rowcount
    def execute(self, sql, params=()):
        if sql.lstrip().startswith('CREATE TABLE IF NOT EXISTS'):
            sql=sql.replace('CREATE TABLE IF NOT EXISTS','CREATE TEMPORARY TABLE IF NOT EXISTS',1)
        if any(table not in {'signed_fiche_documents','signed_fiche_events', 'users', 'permanent_deletion_events'} for table in re.findall(r'(?:FROM|INTO|UPDATE|EXISTS)\s+([a-z_]+)',sql,re.I)):
            raise AssertionError('Unexpected fixture table')
        return self.raw.execute(sql,params)
    def fetchone(self): return self.raw.fetchone()
    def fetchall(self): return self.raw.fetchall()


class TemporaryConnection:
    def __init__(self, db): self.db=db
    def cursor(self): return TemporaryCursor(self.db)
    def commit(self): self.db.commit()
    def rollback(self): self.db.rollback()
    def close(self): pass


@unittest.skipUnless(os.getenv('PRESENCE_MYSQL_TEMP_TEST')=='1','opt-in local MySQL TEMPORARY fixtures')
class SignedFicheMySQLTests(unittest.TestCase):
    def test_real_schema_callback_lease_retry_and_atomic_archive(self):
        config=json.loads((Path(__file__).parent/'.local-test/rebuild/runtime.json').read_text(encoding='utf-8-sig'))
        self.assertEqual(config['database'],'presence_local_rebuild')
        with closing(pymysql.connect(host='127.0.0.1',port=3307,user='root',password=config['db_password'],database=config['database'],
                cursorclass=pymysql.cursors.DictCursor,autocommit=False)) as db, tempfile.TemporaryDirectory() as root:
            proxy=TemporaryConnection(db)
            with proxy.cursor() as cursor:
                ensure_signed_fiche_schema(cursor)
                # Simulate the deployed pre-attribution schema without touching real tables.
                cursor.execute('ALTER TABLE signed_fiche_documents DROP COLUMN manager_username')
                cursor.execute('ALTER TABLE signed_fiche_documents DROP COLUMN is_archived')
                cursor.execute('''INSERT INTO signed_fiche_documents
                    (id,reference_hash,namespace,employee_username,sender_username,employee_name,fiche_month,fiche_year,mode)
                    VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s)''',
                    ('legacy', 'a' * 64, 'fixture', 'synthetic', 'synthetic-admin', 'Legacy fixture', 9, 2026, 'sandbox-only'))
                ensure_signed_fiche_schema(cursor)
                ensure_signed_fiche_schema(cursor)
                cursor.execute("SELECT * FROM signed_fiche_documents WHERE id='legacy'")
                legacy = cursor.fetchone()
                self.assertIsNone(legacy['manager_username'])
                self.assertEqual(legacy['is_archived'], 0)
                self.assertEqual(legacy['reference_hash'], 'a' * 64)
                self.assertEqual(legacy['employee_name'], 'Legacy fixture')
            store=SignedFicheStore(lambda:proxy)
            identifier,reference=store.prepare(employee='synthetic',sender='synthetic-admin',name='Synthetic SQL fixture',
                manager='synthetic-manager',month=9,year=2026,namespace='fixture',mode='sandbox-only')
            store.accepted(identifier,'synthetic_document')
            payload=dict(reference_id=reference,provider='clawshow_esign',namespace='fixture',document_id='synthetic_document',
                event='document.completed',all_signed=True,signed_pdf_url='https://esign.clawshow.ai/esign/synthetic_document/signed.pdf')
            store.receive(payload,validate_signed_url)
            store.receive(payload,validate_signed_url)
            first=store.claim(); self.assertIsNotNone(first); self.assertIsNone(store.claim())
            store.failed(first); self.assertIsNone(store.claim())
            with proxy.cursor() as cursor:
                cursor.execute('UPDATE signed_fiche_documents SET next_attempt_at=UTC_TIMESTAMP(6) WHERE id=%s',(identifier,))
            db.commit()
            row=store.claim(); self.assertIsNotNone(row)
            writer=PdfWriter(); writer.add_blank_page(width=100,height=100); output=io.BytesIO(); writer.write(output)
            service=SignedFiches(lambda:proxy,root); service.save(row,output.getvalue())
            with proxy.cursor() as cursor:
                cursor.execute('SELECT * FROM signed_fiche_documents WHERE id=%s',(identifier,)); final=cursor.fetchone()
            self.assertEqual(final['manager_username'],'synthetic-manager')
            self.assertEqual(final['status'],'ready'); self.assertEqual(final['sha256'],hashlib.sha256(output.getvalue()).hexdigest())
            self.assertEqual(service.read(final),output.getvalue())
            self.assertEqual(service.path(final).parent.relative_to(root).as_posix(),'2026/09')
            store.receive({**payload,'event':'signer.signed','all_signed':False},validate_signed_url)
            with proxy.cursor() as cursor:
                cursor.execute('SELECT status FROM signed_fiche_documents WHERE id=%s',(identifier,))
                self.assertEqual(cursor.fetchone()['status'],'ready')
            # Exercise live role re-read, FOR UPDATE and the independent deletion
            # ledger on real SQL, still shadowing every table on this connection.
            with proxy.cursor() as cursor:
                cursor.execute('''CREATE TEMPORARY TABLE users (username VARCHAR(255) PRIMARY KEY,
                    role_tags TEXT,is_active BOOLEAN DEFAULT 1,session_version INT DEFAULT 0,login_deleted_at DATETIME NULL)''')
                cursor.execute('''CREATE TEMPORARY TABLE permanent_deletion_events (id BIGINT AUTO_INCREMENT PRIMARY KEY,
                    actor_username VARCHAR(255),resource_type VARCHAR(32),resource_id VARCHAR(255),metadata TEXT)''')
                cursor.execute('INSERT INTO users(username,role_tags) VALUES(%s,%s)', ('synthetic-root', '["SuperAdmin","Admin"]'))
            db.commit()
            actor = dict(username='synthetic-root', role_tags=['SuperAdmin','Admin'], session_version=0)
            def load_user(cursor, username):
                cursor.execute('SELECT * FROM users WHERE username=%s', (username,))
                user = cursor.fetchone()
                if user:
                    user['role_tags'] = json.loads(user['role_tags'])
                return user
            app = FastAPI()
            register_signed_fiche_routes(app, get_db_connection=lambda: proxy, get_current_user=lambda: actor,
                                         load_user=load_user, root=root)
            client = TestClient(app)  # No lifespan: the fixture owns all tables.
            try:
                for action, previous in (('archive', False), ('restore', True)):
                    result = client.post('/signed-fiches/batch', json={'ids': [identifier],
                        'action': action, 'expected_archived': previous})
                    self.assertEqual(result.status_code, 200, result.text)
                    self.assertEqual(service.read(final), output.getvalue())
                result = client.post('/signed-fiches/batch', json={'ids': [identifier],
                    'action': 'delete', 'expected_archived': False, 'confirmation': 'SUPPRIMER'})
                self.assertEqual(result.status_code, 200, result.text)
                self.assertFalse(result.json()['cleanup_pending'])
                self.assertFalse(service.path(final).exists())
                self.assertEqual(client.get(f'/signed-fiches/{identifier}/file').status_code, 404)
            finally:
                client.close()
            with proxy.cursor() as cursor:
                cursor.execute('SELECT * FROM permanent_deletion_events WHERE resource_id=%s', (identifier,))
                audit = cursor.fetchone()
                self.assertEqual(audit['resource_type'], 'signed_fiche')
                self.assertEqual(audit['actor_username'], 'synthetic-root')
