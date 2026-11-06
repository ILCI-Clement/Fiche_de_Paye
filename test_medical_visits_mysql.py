"""Opt-in MySQL schema/FK/HTTP verification in a new disposable local database."""
import base64
from io import BytesIO
import json
import os
from pathlib import Path
import secrets
import tempfile
import unittest
from unittest.mock import patch
from zipfile import ZipFile

from fastapi import FastAPI
from fastapi.testclient import TestClient
import pymysql
from pypdf import PdfWriter
from backend.medical_visits import ensure_medical_schema, register_medical_routes
from backend.personnel_sensitive import ensure_local_nir_key_file
from backend.destructive_audit import ensure_destructive_audit_schema
from test_backend_helpers import main as policy_main


@unittest.skipUnless(os.getenv('PRESENCE_MYSQL_TEMP_TEST') == '1', 'opt-in loopback MySQL disposable database')
class MedicalMySQLTests(unittest.TestCase):
    def test_real_schema_foreign_key_and_encrypted_http_roundtrip(self):
        state = json.loads((Path(__file__).parent / '.local-test/rebuild/runtime.json').read_text(encoding='utf-8-sig'))
        self.assertEqual(state['database'], 'presence_local_rebuild')
        options = dict(host='127.0.0.1', port=3307, user='root', password=state['db_password'], cursorclass=pymysql.cursors.DictCursor)
        name = 'medical_test_' + secrets.token_hex(8)
        admin = pymysql.connect(**options, autocommit=True)
        self.addCleanup(admin.close)
        with admin.cursor() as c: c.execute(f'CREATE DATABASE `{name}` CHARACTER SET utf8mb4')
        def cleanup():
            assert name.startswith('medical_test_') and len(name)==29
            with admin.cursor() as c: c.execute(f'DROP DATABASE `{name}`')
        self.addCleanup(cleanup)
        options['database'] = name
        def connection(): return pymysql.connect(**options)
        db = connection(); self.addCleanup(db.close)
        with db.cursor() as c:
            c.execute('''CREATE TABLE users(username VARCHAR(255) PRIMARY KEY,role_tags TEXT,
                is_active INT DEFAULT 1,login_deleted_at TIMESTAMP NULL,session_version INT DEFAULT 0,
                manager_username VARCHAR(255) NULL) ENGINE=InnoDB''')
            c.execute('''CREATE TABLE personnel_register(employee_id BIGINT PRIMARY KEY,full_name VARCHAR(200),
                linked_username VARCHAR(255),is_archived INT DEFAULT 0) ENGINE=InnoDB''')
            c.execute('CREATE TABLE organization_groups(id INT PRIMARY KEY,name VARCHAR(100),is_active INT)')
            c.execute('CREATE TABLE user_group_memberships(username VARCHAR(255),group_id INT,relation VARCHAR(20))')
            c.execute('INSERT INTO users(username,role_tags) VALUES(%s,%s),(%s,%s)',('one','["Employe"]','two','["Responsable"]'))
            c.execute("INSERT INTO personnel_register VALUES(1,'Synthetic Employee','one',0),(2,'Synthetic Manager','two',0)")
            c.execute('INSERT INTO users(username,role_tags) VALUES(%s,%s),(%s,%s),(%s,%s)',
                      ('admin','["Admin"]','super','["SuperAdmin"]','newboss','["Responsable"]'))
            ensure_destructive_audit_schema(c)
        db.commit()
        ensure_medical_schema(connection)
        ensure_medical_schema(connection)
        folder=tempfile.TemporaryDirectory(); self.addCleanup(folder.cleanup)
        keyfile=Path(folder.name)/'keys.json'; ensure_local_nir_key_file(keyfile)
        with patch.dict(os.environ,{'PRESENCE_MEDICAL_KEYS_FILE':str(keyfile),'PRESENCE_MEDICAL_DIR':str(Path(folder.name)/'files')}):
            actor={'username':'one','session_version':0}
            def load(c, name):
                c.execute('SELECT * FROM users WHERE username=%s',(name,))
                row=c.fetchone()
                if row:
                    row['role_tags']=json.loads(row['role_tags'])
                    row['group_ids']=[]; row['managed_group_ids']=[]
                return row
            app=FastAPI(); register_medical_routes(app,get_db_connection=connection,get_current_user=lambda:actor,load_user=load,can_manage=policy_main.user_can_manage)
            with TestClient(app,raise_server_exceptions=False) as client:
                writer=PdfWriter(); writer.add_blank_page(100,100); stream=BytesIO(); writer.write(stream)
                response=client.post('/medical-visits',json={'employee_id':1,'visit_date':'2026-01-01','pdf':{'original_filename':'synthetic.pdf','file_b64':base64.b64encode(stream.getvalue()).decode()}})
                self.assertEqual(response.status_code,200,response.text)
                visit_id=response.json()['id']
                self.assertEqual(client.get('/medical-visits/employees').json()['total'],1)
                self.assertEqual(client.get('/medical-visits/employees/1').json()['total'],1)
                filtered=client.get('/medical-visits/employees?year=2026&month=1')
                self.assertEqual(filtered.status_code,200,filtered.text)
                self.assertEqual(filtered.json()['employees'][0]['latest_id'],visit_id)
                self.assertEqual(filtered.json()['export_count'],1)
                self.assertEqual(client.get('/medical-visits/employees?year=2026&month=2').json()['total'],0)
                exported=client.post('/medical-visits/export',json={'year':2026})
                self.assertEqual(exported.status_code,200,exported.text)
                with ZipFile(BytesIO(exported.content)) as archive:
                    self.assertEqual(len(archive.namelist()),1)
                    self.assertTrue(archive.namelist()[0].startswith('2026/01/'))
                    self.assertEqual(archive.read(archive.namelist()[0]),stream.getvalue())
                self.assertEqual(client.get(f'/medical-visits/{visit_id}/file').content,stream.getvalue())
                actor['username']='two'
                self.assertEqual(client.get(f'/medical-visits/{visit_id}/file').status_code,404)
                with db.cursor() as c:
                    with self.assertRaises(pymysql.err.IntegrityError):
                        c.execute('DELETE FROM personnel_register WHERE employee_id=1')
                db.rollback()
                actor['username']='one'
                corrected=client.put(f'/medical-visits/{visit_id}',json={'expected_version':2,'visit_date':'2026-02-01',
                    'pdf':{'original_filename':'corrected.pdf','file_b64':base64.b64encode(stream.getvalue()).decode()}})
                self.assertEqual(corrected.status_code,200,corrected.text)
                self.assertEqual(client.get(f'/medical-visits/{visit_id}/versions/2/file').content,stream.getvalue())
                self.assertEqual(client.put(f'/medical-visits/{visit_id}',json={'expected_version':2,'visit_date':'2026-03-01'}).status_code,409)
                with db.cursor() as c: c.execute("UPDATE users SET manager_username='two' WHERE username='one'")
                db.commit(); actor['username']='two'
                self.assertEqual(client.get(f'/medical-visits/{visit_id}/versions/2/file').status_code,200)
                self.assertEqual(client.put(f'/medical-visits/{visit_id}',json={'expected_version':3,'visit_date':'2026-03-01'}).status_code,403)
                with db.cursor() as c: c.execute("UPDATE users SET manager_username='newboss' WHERE username='one'")
                db.commit()
                self.assertEqual(client.get(f'/medical-visits/{visit_id}/versions/2/file').status_code,404)
                actor['username']='newboss'
                self.assertEqual(client.get(f'/medical-visits/{visit_id}/versions/2/file').status_code,200)
                actor['username']='admin'
                archived=client.patch(f'/medical-visits/{visit_id}/archive',json={'expected_version':3,'is_archived':True})
                self.assertEqual(archived.status_code,200,archived.text)
                self.assertEqual(client.get('/medical-visits/employees?year=2026').json()['total'],0)
                self.assertEqual(client.get('/medical-visits/employees?year=2026&archived=true').json()['total'],1)
                deletion={'expected_version':4,'confirmation':'SUPPRIMER'}
                self.assertEqual(client.request('DELETE',f'/medical-visits/{visit_id}',json=deletion).status_code,403)
                actor['username']='super'
                result=client.request('DELETE',f'/medical-visits/{visit_id}',json=deletion)
                self.assertEqual(result.status_code,200,result.text)
                self.assertFalse(result.json()['cleanup_pending'])
                self.assertEqual(list((Path(folder.name)/'files').iterdir()),[])
                self.assertEqual(client.get(f'/medical-visits/{visit_id}/versions/2/file').status_code,404)
