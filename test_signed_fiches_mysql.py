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
from pypdf import PdfWriter
from backend.signed_fiche_store import SignedFicheStore, ensure_signed_fiche_schema
from backend.signed_fiches import SignedFiches, validate_signed_url


class TemporaryCursor:
    def __init__(self, db): self.raw=db.cursor()
    def __enter__(self): return self
    def __exit__(self,*_): self.raw.close()
    @property
    def rowcount(self): return self.raw.rowcount
    def execute(self, sql, params=()):
        if sql.lstrip().startswith('CREATE TABLE IF NOT EXISTS'):
            sql=sql.replace('CREATE TABLE IF NOT EXISTS','CREATE TEMPORARY TABLE IF NOT EXISTS',1)
        if any(table not in {'signed_fiche_documents','signed_fiche_events'} for table in re.findall(r'(?:FROM|INTO|UPDATE|EXISTS)\s+([a-z_]+)',sql,re.I)):
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
            with proxy.cursor() as cursor: ensure_signed_fiche_schema(cursor)
            store=SignedFicheStore(lambda:proxy)
            identifier,reference=store.prepare(employee='synthetic',sender='synthetic-admin',name='Synthetic SQL fixture',
                month=9,year=2026,namespace='fixture',mode='sandbox-only')
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
            self.assertEqual(final['status'],'ready'); self.assertEqual(final['sha256'],hashlib.sha256(output.getvalue()).hexdigest())
            self.assertEqual(service.read(final),output.getvalue())
            self.assertEqual(service.path(final).parent.relative_to(root).as_posix(),'2026/09')
            store.receive({**payload,'event':'signer.signed','all_signed':False},validate_signed_url)
            with proxy.cursor() as cursor:
                cursor.execute('SELECT status FROM signed_fiche_documents WHERE id=%s',(identifier,))
                self.assertEqual(cursor.fetchone()['status'],'ready')
