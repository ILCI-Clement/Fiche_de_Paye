"""Opt-in collation regression with connection-local, synthetic TEMPORARY tables."""
import json
import os
from pathlib import Path
import unittest

import pymysql
from backend.signed_fiche_store import signed_fiche_scope


@unittest.skipUnless(os.getenv('PRESENCE_MYSQL_TEMP_TEST') == '1', 'opt-in local MySQL temporary fixtures')
class WorkCenterMySQLTests(unittest.TestCase):
    def test_mixed_collations_do_not_expand_historical_signature_scope(self):
        config = json.loads((Path(__file__).parent / '.local-test/rebuild/runtime.json').read_text(encoding='utf-8-sig'))
        self.assertEqual(config['database'], 'presence_local_rebuild')
        db = pymysql.connect(host='127.0.0.1', port=3307, user='root', password=config['db_password'],
            database=config['database'], cursorclass=pymysql.cursors.DictCursor)
        try:
            with db.cursor() as cursor:
                cursor.execute('''CREATE TEMPORARY TABLE users (username VARCHAR(255),manager_username VARCHAR(255))
                    CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci''')
                cursor.execute('''CREATE TEMPORARY TABLE signed_fiche_documents
                    (id INT,employee_username VARCHAR(255),manager_username VARCHAR(255))
                    CHARACTER SET utf8mb4 COLLATE utf8mb4_general_ci''')
                cursor.execute("INSERT INTO users VALUES ('synthetic-person','synthetic-manager')")
                cursor.executemany('INSERT INTO signed_fiche_documents VALUES (%s,%s,%s)', [
                    (1, 'synthetic-person', 'synthetic-manager'),
                    (2, 'synthetic-person', 'old-manager'),
                    (3, 'SYNTHETIC-PERSON', 'synthetic-manager'),
                    (4, 'synthetic-person', None)])
                scope, params = signed_fiche_scope({'username': 'synthetic-manager', 'role_tags': ['Responsable']})
                cursor.execute(f'SELECT id FROM signed_fiche_documents WHERE {scope}', params)
                self.assertEqual([row['id'] for row in cursor.fetchall()], [1])
                cursor.execute("UPDATE users SET manager_username='new-manager'")
                cursor.execute(f'SELECT id FROM signed_fiche_documents WHERE {scope}', params)
                self.assertEqual(cursor.fetchall(), ())
        finally:
            db.close()


if __name__ == '__main__': unittest.main()
