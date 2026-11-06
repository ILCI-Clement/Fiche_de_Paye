"""Synthetic route, privacy and encrypted-file regression tests."""
import base64
from io import BytesIO
import json
from pathlib import Path
import sqlite3
import tempfile
import unittest
from unittest.mock import patch
from zipfile import ZipFile

from fastapi import FastAPI
from fastapi.testclient import TestClient
from pypdf import PdfWriter
from backend.medical_visits import register_medical_routes
from backend.medical_crypto import protect, reveal
from backend.personnel_sensitive import ensure_local_nir_key_file
from test_management_reports import Connection, Cursor
from test_backend_helpers import main as policy_main


class MedicalCursor(Cursor):
    @property
    def lastrowid(self): return self.cursor.lastrowid


class MedicalConnection(Connection):
    def cursor(self): return MedicalCursor(self.db)


class MedicalVisitsTests(unittest.TestCase):
    def setUp(self):
        folder = tempfile.TemporaryDirectory()
        self.addCleanup(folder.cleanup)
        self.folder = Path(folder.name)
        keyfile = self.folder / 'keys.json'
        ensure_local_nir_key_file(keyfile)
        env = patch.dict('os.environ', {'PRESENCE_MEDICAL_KEYS_FILE': str(keyfile), 'PRESENCE_MEDICAL_DIR': str(self.folder / 'pdf')})
        env.start(); self.addCleanup(env.stop)
        self.db = sqlite3.connect(':memory:', check_same_thread=False)
        self.db.row_factory = sqlite3.Row
        self.addCleanup(self.db.close)
        self.db.executescript('''
            CREATE TABLE users(username TEXT PRIMARY KEY,role_tags TEXT,is_active INT,login_deleted_at TEXT,session_version INT,manager_username TEXT);
            CREATE TABLE personnel_register(employee_id INTEGER PRIMARY KEY,full_name TEXT,linked_username TEXT,is_archived INT);
            CREATE TABLE organization_groups(id INTEGER PRIMARY KEY,name TEXT,is_active INT);
            CREATE TABLE user_group_memberships(username TEXT,group_id INT,relation TEXT);
            CREATE TABLE medical_visits(id INTEGER PRIMARY KEY,employee_id INT,visit_date TEXT,
                original_filename TEXT,stored_filename TEXT,size_bytes INT DEFAULT 0,version INT DEFAULT 1,
                created_by TEXT,created_at TEXT DEFAULT CURRENT_TIMESTAMP,is_archived INT DEFAULT 0);
            CREATE TABLE medical_visit_events(id INTEGER PRIMARY KEY,visit_id INT,actor_username TEXT,action TEXT,metadata TEXT);
            CREATE TABLE medical_visit_versions(id INTEGER PRIMARY KEY,visit_id INT,version INT,visit_date TEXT,
                original_filename TEXT,stored_filename TEXT,size_bytes INT,changed_by TEXT,created_at TEXT DEFAULT CURRENT_TIMESTAMP,
                UNIQUE(visit_id,version));
            CREATE TABLE permanent_deletion_events(id INTEGER PRIMARY KEY,actor_username TEXT,resource_type TEXT,resource_id TEXT,metadata TEXT);
            INSERT INTO personnel_register VALUES(1,'Employee One','one',0),(2,'Employee Two','two',0),
                (3,'Archived Employee','old',1),(4,'Hidden Super','super',0),(5,'No login',NULL,0);
            INSERT INTO organization_groups VALUES(1,'Research',1);
            INSERT INTO user_group_memberships VALUES('one',1,'member');
        ''')
        for username, roles in [('admin',['Admin']),('super',['SuperAdmin']),('one',['Employe']),('two',['Employe']),('boss',['Responsable']),('old',['Employe'])]:
            self.db.execute('INSERT INTO users VALUES(?,?,1,NULL,0,NULL)', (username,json.dumps(roles)))
        self.db.commit()
        self.actor = {'username': 'one', 'session_version': 0}
        def load(c, username):
            c.execute('SELECT * FROM users WHERE username=%s', (username,))
            row = c.fetchone()
            if row:
                row['role_tags'] = json.loads(row['role_tags'])
                c.execute('SELECT group_id,relation FROM user_group_memberships WHERE username=%s', (username,))
                memberships = c.fetchall()
                row['group_ids'] = [m['group_id'] for m in memberships if m['relation']=='member']
                row['managed_group_ids'] = [m['group_id'] for m in memberships if m['relation']=='manager']
            return row
        app = FastAPI()
        register_medical_routes(app, get_db_connection=lambda:MedicalConnection(self.db), get_current_user=lambda:self.actor, load_user=load, can_manage=policy_main.user_can_manage)
        self.client = TestClient(app, raise_server_exceptions=False)
        self.addCleanup(self.client.close)
        writer = PdfWriter(); writer.add_blank_page(100,100)
        stream = BytesIO(); writer.write(stream)
        self.content = stream.getvalue()
        self.pdf = {'original_filename':'attestation.pdf','file_b64':base64.b64encode(self.content).decode()}

    def create(self, employee_id=1, pdf=None, date='2026-01-15'):
        return self.client.post('/medical-visits', json={'employee_id':employee_id,'visit_date':date,'pdf':pdf})

    def test_zip_contains_only_named_originals_and_rejects_unauthorized_selection(self):
        first = self.create(pdf=self.pdf).json()['id']
        second = self.create(pdf=self.pdf,date='2026-02-12').json()['id']
        empty = self.create(date='2026-03-10').json()['id']
        payload = {'year':2026,'format':'zip','visit_ids':[first,second]}
        result = self.client.post('/medical-visits/export',json=payload)
        self.assertEqual(result.status_code,200,result.text[:100])
        with ZipFile(BytesIO(result.content)) as archive:
            self.assertEqual(len(archive.namelist()),2)
            for path in archive.namelist():
                self.assertTrue(path.startswith('2026/'))
                self.assertIn('Employee One',path)
                self.assertNotIn('..',path)
                self.assertEqual(archive.read(path),self.content)
            self.assertNotIn('Index.pdf',archive.namelist())
        self.actor['username']='two'
        self.assertEqual(self.client.post('/medical-visits/export',json=payload).status_code,404)
        self.actor['username']='one'
        self.assertEqual(self.client.post('/medical-visits/export',json={**payload,'visit_ids':[first,empty]}).status_code,404)

    def test_personal_export_includes_all_years_and_archives_but_not_other_people(self):
        first = self.create(pdf=self.pdf,date='2024-02-01').json()['id']
        self.create(pdf=self.pdf,date='2026-01-01')
        self.db.execute('UPDATE medical_visits SET is_archived=1 WHERE id=?',(first,))
        self.db.commit()
        self.actor['username']='admin'
        self.create(employee_id=2,pdf=self.pdf,date='2025-02-01')
        result = self.client.post('/medical-visits/export',json={'year':2026,'format':'zip','employee_id':1,'all_history':True})
        self.assertEqual(result.status_code,200,result.text[:100])
        with ZipFile(BytesIO(result.content)) as archive:
            self.assertEqual(len(archive.namelist()),2)
            self.assertTrue(any(p.startswith('2024/') for p in archive.namelist()))
            self.assertFalse(any('Employee_Two' in p for p in archive.namelist()))
        self.assertEqual(self.client.post('/medical-visits/export',json={'year':2026,'format':'zip','all_history':True}).status_code,400)
        self.actor['username']='two'
        self.assertEqual(self.client.post('/medical-visits/export',json={'year':2026,'format':'zip','employee_id':1,'all_history':True}).status_code,404)

    def test_date_first_later_pdf_encrypted_and_audited(self):
        result = self.create()
        self.assertEqual(result.status_code,200,result.text)
        identity = result.json()['id']
        self.assertEqual(self.client.get(f'/medical-visits/{identity}/file').status_code,404)
        self.assertEqual(self.client.post(f'/medical-visits/{identity}/file',json={**self.pdf,'expected_version':1}).status_code,200)
        self.assertEqual(self.client.post(f'/medical-visits/{identity}/file',json={**self.pdf,'expected_version':2}).status_code,409)
        stored = next((self.folder/'pdf').iterdir()).read_bytes()
        self.assertNotIn(b'%PDF-',stored)
        response = self.client.get(f'/medical-visits/{identity}/file')
        self.assertEqual(response.content,self.content)
        self.assertIn('no-store',response.headers['cache-control'])
        self.assertEqual(response.headers['x-content-type-options'],'nosniff')
        self.assertEqual([r[0] for r in self.db.execute('SELECT action FROM medical_visit_events')],['create','upload','read'])

    def test_employee_and_responsable_cannot_access_other_employees(self):
        identity = self.create(pdf=self.pdf).json()['id']
        for name in ['two','boss']:
            self.actor['username'] = name
            self.assertEqual(self.client.get('/medical-visits/employees/1').status_code,404)
            self.assertEqual(self.client.get(f'/medical-visits/{identity}/file').status_code,404)
            self.assertEqual(self.client.post(f'/medical-visits/{identity}/file',json={**self.pdf,'expected_version':2}).status_code,404)
            self.assertEqual(self.create().status_code,404)
        self.actor['username']='two'
        people = self.client.get('/medical-visits/employees').json()['employees']
        self.assertEqual([p['employee_id'] for p in people],[2])
        self.actor['username']='boss'
        self.assertEqual(self.client.get('/medical-visits/employees').json()['total'],0)

    def test_admin_scope_hidden_super_archives_and_unlinked(self):
        self.actor['username']='admin'
        response=self.client.get('/medical-visits/employees')
        ids={p['employee_id'] for p in response.json()['employees']}
        self.assertEqual(ids,{1,2,5})
        self.assertEqual(self.client.get('/medical-visits/employees/4').status_code,404)
        self.assertEqual(self.create(4).status_code,404)
        self.assertEqual(self.create(5).status_code,200)
        archived=self.client.get('/medical-visits/employees?archived=true').json()
        self.assertEqual([p['employee_id'] for p in archived['employees']],[3])
        self.assertEqual(self.create(3).status_code,409)
        self.actor['username']='super'
        self.assertIn(4,{p['employee_id'] for p in self.client.get('/medical-visits/employees').json()['employees']})

    def test_latest_actual_date_history_search_and_data_minimization(self):
        self.create(date='2026-02-15'); self.create(date='2026-01-15')
        response=self.client.get('/medical-visits/employees?q=Employee')
        person=response.json()['employees'][0]
        self.assertEqual(person['visit_date'],'2026-02-15')
        self.assertEqual(person['visit_count'],2)
        self.assertEqual(person['departments'],['Research'])
        self.assertNotIn('stored_filename',response.text)
        self.assertNotIn('email',response.text)
        history=self.client.get('/medical-visits/employees/1').json()
        self.assertEqual([v['visit_date'] for v in history['visits']],['2026-02-15','2026-01-15'])
        self.assertEqual(self.client.get('/medical-visits/employees/1?page=2').json()['visits'],[])

    def test_invalid_pdf_future_date_and_missing_keys_are_atomic(self):
        for pdf in [{**self.pdf,'file_b64':'%%%INVALID%%%'}, {**self.pdf,'file_b64':base64.b64encode(b'%PDF-fake').decode()}, {**self.pdf,'original_filename':'file.html'}]:
            response=self.create(pdf=pdf)
            self.assertEqual(response.status_code,400,response.text)
        self.assertEqual(self.create(date='2999-01-01').status_code,400)
        self.assertEqual(self.create(date='Invalid Date').status_code,422)
        with patch.dict('os.environ',{'PRESENCE_MEDICAL_KEYS_FILE':''}):
            self.assertEqual(self.create(pdf=self.pdf).status_code,503)
            self.assertEqual(self.create().status_code,200)
        self.assertEqual(self.db.execute('SELECT COUNT(*) FROM medical_visits').fetchone()[0],1)
        self.assertEqual(self.db.execute('SELECT COUNT(*) FROM medical_visit_events').fetchone()[0],1)

    def test_revoked_account_or_session_cannot_read_or_write(self):
        identity=self.create(pdf=self.pdf).json()['id']
        self.db.execute("UPDATE users SET session_version=1 WHERE username='one'"); self.db.commit()
        for path in ['/medical-visits/employees','/medical-visits/employees/1',f'/medical-visits/{identity}/file']:
            self.assertEqual(self.client.get(path).status_code,401)
        self.assertEqual(self.create().status_code,401)

    def test_disabled_logins_leave_default_list_and_link_change_rechecks_ownership(self):
        identity=self.create(pdf=self.pdf).json()['id']
        self.db.execute("UPDATE users SET is_active=0 WHERE username='two'"); self.db.commit()
        self.actor['username']='admin'
        self.assertNotIn(2,{p['employee_id'] for p in self.client.get('/medical-visits/employees').json()['employees']})
        self.db.execute("UPDATE personnel_register SET linked_username='admin' WHERE employee_id=1"); self.db.commit()
        self.actor['username']='one'
        self.assertEqual(self.client.get(f'/medical-visits/{identity}/file').status_code,404)
        self.assertEqual(self.client.get('/medical-visits/employees').json()['total'],0)

    def test_audit_failure_leaves_no_pdf_or_row(self):
        self.db.execute("CREATE TRIGGER fail_audit BEFORE INSERT ON medical_visit_events WHEN NEW.action='upload' BEGIN SELECT RAISE(ABORT,'synthetic'); END")
        response=self.create(pdf=self.pdf)
        self.assertEqual(response.status_code,500)
        self.assertEqual(list((self.folder/'pdf').glob('*')),[])
        self.assertEqual(self.db.execute('SELECT COUNT(*) FROM medical_visits').fetchone()[0],0)

    def test_envelope_binds_identity_and_detects_tampering(self):
        envelope=protect(self.content,'employee:one')
        self.assertEqual(reveal(envelope,'employee:one'),self.content)
        with self.assertRaises(Exception): reveal(envelope,'employee:two')
        with self.assertRaises(Exception): reveal(envelope[:-1]+bytes([envelope[-1]^1]),'employee:one')

    def test_year_month_filters_and_zip_follow_visit_date_not_upload_date(self):
        first=self.create(pdf=self.pdf,date='2026-01-15').json()['id']
        self.create(pdf=self.pdf,date='2026-02-20')
        self.create(pdf=self.pdf,date='2025-01-15')
        self.create(date='2026-03-10')
        response=self.client.get('/medical-visits/employees?year=2026')
        self.assertEqual(response.status_code,200,response.text)
        person=response.json()['employees'][0]
        self.assertEqual(person['visit_count'],3)
        self.assertEqual(person['visit_date'],'2026-03-10')
        self.assertEqual(response.json()['export_count'],2)
        self.assertEqual(response.json()['years'],[2026,2025])
        month=self.client.get('/medical-visits/employees?year=2026&month=1').json()
        self.assertEqual(month['employees'][0]['latest_id'],first)
        self.assertEqual(month['employees'][0]['visit_count'],1)
        self.assertEqual(self.client.get('/medical-visits/employees?year=2026&month=4').json()['total'],0)
        history=self.client.get('/medical-visits/employees/1?year=2026&month=1').json()
        self.assertEqual([v['id'] for v in history['visits']],[first])
        result=self.client.post('/medical-visits/export',json={'year':2026})
        self.assertEqual(result.status_code,200,result.text)
        self.assertIn('no-store',result.headers['cache-control'])
        with ZipFile(BytesIO(result.content)) as archive:
            names=archive.namelist()
            self.assertEqual(len(names),2)
            self.assertTrue(names[0].startswith('2026/01/Employee One_1_'))
            self.assertTrue(names[1].startswith('2026/02/Employee One_1_'))
            self.assertEqual(archive.read(names[0]),self.content)
        monthly=self.client.post('/medical-visits/export',json={'year':2025,'month':1})
        self.assertEqual(len(ZipFile(BytesIO(monthly.content)).namelist()),1)
        self.assertEqual(self.client.get('/medical-visits/employees?month=1').status_code,400)
        self.assertEqual(self.client.post('/medical-visits/export',json={'year':2026,'month':13}).status_code,422)

    def test_export_enforces_owner_and_admin_search_archive_scope(self):
        self.create(pdf=self.pdf)
        self.actor['username']='two'; self.create(2,pdf=self.pdf)
        own=self.client.post('/medical-visits/export',json={'year':2026})
        with ZipFile(BytesIO(own.content)) as archive:
            self.assertTrue(all('/Employee Two_2_' in p for p in archive.namelist()))
        self.assertEqual(self.client.post('/medical-visits/export',json={'year':2026,'employee_id':1}).status_code,404)
        self.actor['username']='boss'
        self.assertEqual(self.client.post('/medical-visits/export',json={'year':2026}).status_code,400)
        self.actor['username']='admin'
        result=self.client.post('/medical-visits/export',json={'year':2026,'q':'One'})
        self.assertEqual(len(ZipFile(BytesIO(result.content)).namelist()),1)
        self.db.execute('UPDATE personnel_register SET is_archived=1 WHERE employee_id=1'); self.db.commit()
        self.assertEqual(self.client.post('/medical-visits/export',json={'year':2026,'q':'One'}).status_code,400)
        archived=self.client.post('/medical-visits/export',json={'year':2026,'archived':True})
        self.assertEqual(len(ZipFile(BytesIO(archived.content)).namelist()),1)

    def test_export_all_pages_unique_paths_and_atomic_failure(self):
        identity=self.create(pdf=self.pdf).json()['id']
        # Same display names and filenames cannot overwrite archive members.
        for number in range(21): self.create(pdf=self.pdf)
        listed=self.client.get('/medical-visits/employees/1?year=2026').json()
        self.assertEqual(len(listed['visits']),20)
        self.assertEqual(listed['total'],22)
        exported=self.client.post('/medical-visits/export',json={'year':2026})
        self.assertEqual(exported.status_code,200,exported.text)
        names=ZipFile(BytesIO(exported.content)).namelist()
        self.assertEqual(len(names),22); self.assertEqual(len(set(names)),22)
        before=self.db.execute("SELECT COUNT(*) FROM medical_visit_events WHERE action='export'").fetchone()[0]
        path=self.db.execute('SELECT stored_filename FROM medical_visits WHERE id=?',(identity,)).fetchone()[0]
        (self.folder/'pdf'/path).unlink()
        self.assertEqual(self.client.post('/medical-visits/export',json={'year':2026}).status_code,503)
        self.assertEqual(self.db.execute("SELECT COUNT(*) FROM medical_visit_events WHERE action='export'").fetchone()[0],before)

    def test_export_file_and_byte_limits_fail_without_partial_audit(self):
        identity=self.create(pdf=self.pdf).json()['id']
        with patch('backend.medical_visits.MAX_ARCHIVE_BYTES',1):
            self.assertEqual(self.client.post('/medical-visits/export',json={'year':2026}).status_code,400)
        self.db.execute('UPDATE medical_visits SET size_bytes=0'); self.db.commit()
        with patch('backend.medical_visits.MAX_ARCHIVE_BYTES',1):
            self.assertEqual(self.client.post('/medical-visits/export',json={'year':2026}).status_code,400)
        self.assertEqual(self.db.execute("SELECT COUNT(*) FROM medical_visit_events WHERE action='export'").fetchone()[0],0)
        self.db.executemany('''INSERT INTO medical_visits(employee_id,visit_date,created_by,stored_filename,original_filename,size_bytes)
            SELECT employee_id,visit_date,created_by,stored_filename,original_filename,size_bytes FROM medical_visits WHERE id=?''',[(identity,)]*100)
        self.db.commit()
        self.assertEqual(self.client.post('/medical-visits/export',json={'year':2026}).status_code,400)

    def detail_visit(self, identity):
        return self.client.get(f'/medical-visits/{identity}/versions').json()['current']

    def correction(self, identity, version, pdf=None, day='2026-02-10'):
        return self.client.put(f'/medical-visits/{identity}', json={'expected_version':version,'visit_date':day,'pdf':pdf})

    def test_team_policy_is_live_new_manager_reads_old_history_but_cannot_edit(self):
        identity=self.create(pdf=self.pdf).json()['id']
        self.correction(identity,2,pdf=self.pdf)
        self.db.execute("UPDATE users SET manager_username='boss' WHERE username='one'"); self.db.commit()
        self.actor['username']='boss'
        people=self.client.get('/medical-visits/employees?year=2026').json()
        self.assertEqual([p['employee_id'] for p in people['employees']],[1])
        self.assertFalse(self.client.get('/medical-visits/employees/1').json()['employee']['can_edit'])
        self.assertEqual(self.client.get(f'/medical-visits/{identity}/file').status_code,200)
        self.assertEqual(self.client.get(f'/medical-visits/{identity}/versions/2/file').status_code,200)
        self.assertEqual(self.correction(identity,3).status_code,403)
        self.assertEqual(self.create().status_code,403)
        self.assertEqual(self.client.patch(f'/medical-visits/{identity}/archive',json={'expected_version':3,'is_archived':True}).status_code,403)
        self.db.execute("UPDATE users SET role_tags='[\"Responsable\"]' WHERE username='two'")
        self.db.execute("UPDATE users SET manager_username='two' WHERE username='one'"); self.db.commit()
        self.assertEqual(self.client.get(f'/medical-visits/{identity}/file').status_code,404)
        self.assertEqual(self.client.get(f'/medical-visits/{identity}/versions/2/file').status_code,404)
        self.actor['username']='two'
        self.assertEqual(self.client.get(f'/medical-visits/{identity}/versions/2/file').status_code,200)
        self.assertEqual(self.client.get('/medical-visits/employees?mine=true').json()['employees'][0]['employee_id'],2)

    def test_department_management_reuses_same_policy_without_direct_binding(self):
        identity=self.create(pdf=self.pdf).json()['id']
        self.db.execute("INSERT INTO user_group_memberships VALUES('boss',1,'manager')"); self.db.commit()
        self.actor['username']='boss'
        self.assertEqual(self.client.get(f'/medical-visits/{identity}/file').status_code,200)
        self.assertIn(1,[p['employee_id'] for p in self.client.get('/medical-visits/employees').json()['employees']])
        self.db.execute("DELETE FROM user_group_memberships WHERE username='boss'"); self.db.commit()
        self.assertEqual(self.client.get(f'/medical-visits/{identity}/file').status_code,404)

    def test_owner_correction_preserves_pdf_versions_and_dates_with_conflict_guard(self):
        identity=self.create(pdf=self.pdf).json()['id']
        old=self.detail_visit(identity)
        replacement={**self.pdf,'original_filename':'corrected.pdf'}
        result=self.correction(identity,old['version'],pdf=replacement)
        self.assertEqual(result.status_code,200,result.text)
        current=self.detail_visit(identity)
        self.assertEqual(current['version'],3)
        self.assertEqual(current['visit_date'],'2026-02-10')
        self.assertEqual(current['original_filename'],'corrected.pdf')
        previous=self.client.get(f'/medical-visits/{identity}/versions').json()['versions']
        self.assertEqual(previous[0]['visit_date'],'2026-01-15')
        self.assertEqual(previous[0]['original_filename'],'attestation.pdf')
        self.assertEqual(self.client.get(f'/medical-visits/{identity}/versions/2/file').content,self.content)
        self.assertEqual(len(list((self.folder/'pdf').glob('*'))),2)
        self.assertEqual(self.correction(identity,2).status_code,409)
        self.assertEqual(self.correction(identity,3,day='2999-01-01').status_code,400)
        self.assertEqual(self.correction(identity,3,day='2026-03-01').status_code,200)
        self.assertEqual(self.detail_visit(identity)['original_filename'],'corrected.pdf')
        self.assertEqual(len(list((self.folder/'pdf').glob('*'))),2)
        self.actor['username']='admin'
        self.assertEqual(self.correction(identity,4,day='2026-04-01').status_code,200)

    def test_correction_failure_keeps_original_file_and_history_atomic(self):
        identity=self.create(pdf=self.pdf).json()['id']
        before=list((self.folder/'pdf').iterdir())
        self.db.execute("CREATE TRIGGER fail_correct BEFORE INSERT ON medical_visit_events WHEN NEW.action='correct' BEGIN SELECT RAISE(ABORT,'synthetic'); END")
        self.assertEqual(self.correction(identity,2,pdf=self.pdf).status_code,500)
        self.assertEqual(self.detail_visit(identity)['version'],2)
        self.assertEqual(list((self.folder/'pdf').iterdir()),before)
        self.assertEqual(self.client.get(f'/medical-visits/{identity}/versions').json()['total'],0)

    def test_late_attachment_rejects_stale_date_and_archive_is_read_only(self):
        identity=self.create().json()['id']
        self.assertEqual(self.correction(identity,1,day='2026-02-01').status_code,200)
        self.assertEqual(self.client.post(f'/medical-visits/{identity}/file',json={**self.pdf,'expected_version':1}).status_code,409)
        self.assertEqual(self.client.post(f'/medical-visits/{identity}/file',json={**self.pdf,'expected_version':2}).status_code,200)
        self.actor['username']='admin'
        self.client.patch(f'/medical-visits/{identity}/archive',json={'expected_version':3,'is_archived':True})
        self.actor['username']='one'
        self.assertEqual(self.correction(identity,4).status_code,409)
        self.assertEqual(self.client.get('/medical-visits/employees/1?archived=true').json()['total'],1)
        self.assertEqual(self.client.get(f'/medical-visits/{identity}/file').status_code,200)

    def test_delete_audit_failure_restores_rows_versions_and_files(self):
        identity=self.create(pdf=self.pdf).json()['id']; self.correction(identity,2,pdf=self.pdf)
        self.actor['username']='super'
        self.client.patch(f'/medical-visits/{identity}/archive',json={'expected_version':3,'is_archived':True})
        before=list((self.folder/'pdf').iterdir())
        self.db.execute("CREATE TRIGGER fail_deletion BEFORE INSERT ON permanent_deletion_events BEGIN SELECT RAISE(ABORT,'synthetic'); END")
        result=self.client.request('DELETE',f'/medical-visits/{identity}',json={'expected_version':4,'confirmation':'SUPPRIMER'})
        self.assertEqual(result.status_code,500)
        self.assertEqual(self.detail_visit(identity)['version'],4)
        self.assertEqual(self.client.get(f'/medical-visits/{identity}/versions').json()['total'],1)
        self.assertEqual(list((self.folder/'pdf').iterdir()),before)

    def test_archive_restore_and_super_delete_remove_all_files_keep_minimal_audit(self):
        identity=self.create(pdf=self.pdf).json()['id']
        self.correction(identity,2,pdf=self.pdf)
        self.actor['username']='admin'
        archived=self.client.patch(f'/medical-visits/{identity}/archive',json={'expected_version':3,'is_archived':True})
        self.assertEqual(archived.status_code,200,archived.text)
        self.assertEqual(self.client.get('/medical-visits/employees?year=2026').json()['total'],0)
        self.assertEqual(self.client.get('/medical-visits/employees?year=2026&archived=true').json()['total'],1)
        self.assertEqual(self.client.get('/medical-visits/employees/1?archived=true').json()['total'],1)
        self.assertEqual(self.correction(identity,4).status_code,409)
        delete=lambda version,word: self.client.request('DELETE',f'/medical-visits/{identity}',json={'expected_version':version,'confirmation':word})
        self.assertEqual(delete(4,'SUPPRIMER').status_code,403)
        self.actor['username']='super'
        self.assertEqual(delete(3,'SUPPRIMER').status_code,409)
        self.assertEqual(delete(4,'wrong').status_code,400)
        self.actor['username']='admin'
        self.assertEqual(self.client.patch(f'/medical-visits/{identity}/archive',json={'expected_version':4,'is_archived':False}).status_code,200)
        self.actor['username']='super'
        self.assertEqual(delete(5,'SUPPRIMER').status_code,409)
        self.client.patch(f'/medical-visits/{identity}/archive',json={'expected_version':5,'is_archived':True})
        result=delete(6,'SUPPRIMER')
        self.assertEqual(result.status_code,200,result.text)
        self.assertFalse(result.json()['cleanup_pending'])
        self.assertEqual(list((self.folder/'pdf').iterdir()),[])
        self.assertEqual(self.client.get(f'/medical-visits/{identity}/file').status_code,404)
        self.assertEqual(self.db.execute('SELECT COUNT(*) FROM medical_visit_versions').fetchone()[0],0)
        self.assertEqual(self.db.execute('SELECT COUNT(*) FROM medical_visit_events').fetchone()[0],0)
        ledger=dict(self.db.execute('SELECT * FROM permanent_deletion_events').fetchone())
        self.assertEqual(ledger['resource_type'],'medical_visit')
        self.assertNotIn('Employee One',str(ledger))
        self.assertNotIn('attestation.pdf',str(ledger))
        self.assertEqual(json.loads(ledger['metadata'])['deleted_files'],2)


if __name__ == '__main__': unittest.main()
