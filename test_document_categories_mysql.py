"""Opt-in migration regression; writes are restricted to connection-local tables."""
import json
import os
from pathlib import Path
import unittest

import pymysql
from fastapi import HTTPException
from backend.document_categories import ensure_category_schema, label_key, resolve_category


class TemporaryConnection:
    def __init__(self, db): self.db = db
    def cursor(self): return TemporaryCursor(self.db)
    def commit(self): self.db.commit()
    def rollback(self): self.db.rollback()
    def close(self): pass


class TemporaryCursor:
    def __init__(self, db): self.raw = db.cursor()
    def __enter__(self): return self
    def __exit__(self, *_): self.raw.close()
    def execute(self, query, params=()):
        if query.lstrip().startswith('CREATE TABLE IF NOT EXISTS'):
            query = query.replace('CREATE TABLE IF NOT EXISTS', 'CREATE TEMPORARY TABLE IF NOT EXISTS', 1)
        if any(word in query.upper() for word in ('ALTER TABLE', 'DROP TABLE', 'DELETE FROM')):
            raise AssertionError('Fixture cannot alter persistent schema or remove rows')
        return self.raw.execute(query, params)
    def fetchone(self): return self.raw.fetchone()
    def fetchall(self): return self.raw.fetchall()


@unittest.skipUnless(os.getenv('PRESENCE_MYSQL_TEMP_TEST') == '1', 'opt-in local MySQL temporary fixtures')
class DocumentCategoryMySQLTests(unittest.TestCase):
    def test_repeatable_migration_preserves_legacy_receipts_and_twenty_character_types(self):
        config = json.loads((Path(__file__).parent / '.local-test/rebuild/runtime.json').read_text(encoding='utf-8-sig'))
        self.assertEqual(config['database'], 'presence_local_rebuild')
        db = pymysql.connect(host='127.0.0.1', port=3307, user='root', password=config['db_password'],
            database=config['database'], cursorclass=pymysql.cursors.DictCursor)
        try:
            with db.cursor() as cursor:
                cursor.execute('CREATE TEMPORARY TABLE transport_receipts (id INT,document_type VARCHAR(20) NULL,other_type VARCHAR(120) NULL)')
                cursor.execute("INSERT INTO transport_receipts VALUES (1,NULL,NULL),(2,'autre','Legacy taxi')")
                cursor.execute('SELECT * FROM transport_receipts ORDER BY id'); before = cursor.fetchall()
            proxy = TemporaryConnection(db)
            ensure_category_schema(lambda: proxy)
            ensure_category_schema(lambda: proxy)
            with db.cursor() as cursor:
                cursor.execute('SELECT * FROM transport_receipts ORDER BY id'); self.assertEqual(cursor.fetchall(), before)
                cursor.execute('SELECT code FROM document_categories ORDER BY code')
                self.assertEqual([r['code'] for r in cursor.fetchall()], ['autre', 'transport'])
                code = 'c_' + 'a' * 18
                cursor.execute('INSERT INTO document_categories (code,label,label_key) VALUES (%s,%s,%s)', (code, 'Hôtel', label_key('Hôtel')))
                cursor.execute('INSERT INTO transport_receipts VALUES (3,%s,NULL)', (code,))
                self.assertEqual(resolve_category(cursor, code, None), ('Hôtel', None))
                with self.assertRaises(pymysql.err.IntegrityError):
                    cursor.execute('INSERT INTO document_categories (code,label,label_key) VALUES (%s,%s,%s)', ('c_' + 'b' * 18, 'HOTEL', label_key('HOTEL')))
                cursor.execute('UPDATE document_categories SET is_active=0 WHERE code=%s', (code,))
                with self.assertRaises(HTTPException) as error: resolve_category(cursor, code, None)
                self.assertEqual(error.exception.status_code, 409)
                cursor.execute('SELECT document_type FROM transport_receipts WHERE id=3')
                self.assertEqual(cursor.fetchone()['document_type'], code)
        finally:
            db.close()


if __name__ == '__main__': unittest.main()
