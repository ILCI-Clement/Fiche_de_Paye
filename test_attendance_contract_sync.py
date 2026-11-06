import copy
import sqlite3
import unittest

from backend.attendance_contract_sync import refresh_contract_snapshots
from backend.attendance_revision import config_revision
from backend.attendance_sheet import new_sheet
from backend.leave_store import LeaveStore


class ContractSnapshotTests(unittest.TestCase):
    def setUp(self):
        self.db = sqlite3.connect(':memory:')
        self.db.row_factory = sqlite3.Row
        self.addCleanup(self.db.close)
        self.db.executescript('''
            CREATE TABLE personnel_contracts(id INT,username TEXT,employee_type TEXT,start_date TEXT,end_date TEXT,is_cdi INT,is_archived INT);
            INSERT INTO personnel_contracts VALUES(1,'member','salarie','2026-08-03','2027-08-21',0,0);
            CREATE TABLE signed_fiche_documents(employee_username TEXT,fiche_year INT,fiche_month INT,mode TEXT,status TEXT);
        ''')
        self.store = LeaveStore.borrow(self.db.cursor())
        self.sheet = new_sheet({'username':'member','email':'synthetic@example.com'}, None,
            {'id':1,'username':'member','employee_type':'salarie','start_date':'2025-09-15','end_date':'2026-09-14','is_cdi':0}, 9, 2026, [])
        self.sheet.update(nom='Historical document name', custom={'keep': True},
            calendar_overrides={'2026-09-02': {'matin': ['Absence', 'synthetic note']}})
        self.config = {'employes_data': [self.sheet], 'unknown': {'keep': True}}

    def test_read_check_does_not_write_and_refresh_preserves_attendance_identity_and_history(self):
        before = copy.deepcopy(self.config)
        old_revision = config_revision(before)
        self.assertEqual(refresh_contract_snapshots(self.store, self.config), [self.sheet['id']])
        self.assertEqual(self.config, before)
        refresh_contract_snapshots(self.store, self.config, actor={'username':'operator'})
        for key in ('id','nom','calendar_overrides','planning_detail','custom'):
            self.assertEqual(self.sheet[key], before['employes_data'][0][key])
        self.assertEqual(self.sheet['ddc'], '2026-08-03')
        self.assertEqual(self.sheet['fdc'], '2027-08-21')
        self.assertEqual(self.sheet['sheet_period_end'], '2026-09-30')
        self.assertEqual(self.sheet['contract_sync_history'][0]['before']['contract_snapshot']['end_date'], '2026-09-14')
        self.assertNotEqual(config_revision(self.config), old_revision)
        self.assertEqual(refresh_contract_snapshots(self.store, self.config, actor={'username':'operator'}), [])
        self.assertEqual(len(self.sheet['contract_sync_history']), 1)

    def test_archived_pending_and_completed_sources_are_unchanged(self):
        original = copy.deepcopy(self.config)
        for status in ('pending','sent','partially_signed','completed','archived','uncertain'):
            with self.subTest(status=status):
                self.db.execute('DELETE FROM signed_fiche_documents')
                self.db.execute('INSERT INTO signed_fiche_documents VALUES(?,?,?,?,?)', ('member',2026,9,'production',status))
                self.assertEqual(refresh_contract_snapshots(self.store, self.config, actor={'username':'operator'}), [])
                self.assertEqual(self.config, original)
        self.db.execute('DELETE FROM signed_fiche_documents')
        self.sheet['is_archived'] = True
        before = copy.deepcopy(self.config)
        self.assertEqual(refresh_contract_snapshots(self.store, self.config, actor={'username':'operator'}), [])
        self.assertEqual(self.config, before)

    def test_missing_archived_wrong_identity_and_nonintersecting_contract_are_retained(self):
        for sql in ("DELETE FROM personnel_contracts", "UPDATE personnel_contracts SET is_archived=1",
                    "UPDATE personnel_contracts SET username='other'", "UPDATE personnel_contracts SET start_date='2027-01-01'"):
            with self.subTest(sql=sql):
                self.db.execute('SAVEPOINT variant')
                self.db.execute(sql)
                before = copy.deepcopy(self.config)
                self.assertEqual(refresh_contract_snapshots(self.store, self.config, actor={'username':'operator'}), [])
                self.assertEqual(self.config, before)
                self.db.execute('ROLLBACK TO variant')

    def test_cdi_and_shortened_period_keep_original_attendance(self):
        self.db.execute('UPDATE personnel_contracts SET end_date=NULL,is_cdi=1')
        refresh_contract_snapshots(self.store, self.config, actor={'username':'operator'})
        self.assertTrue(self.sheet['cdi'])
        self.assertEqual(self.sheet['fdc'], 'Pas de fin')
        self.db.execute("UPDATE personnel_contracts SET end_date='2026-09-10',is_cdi=0")
        refresh_contract_snapshots(self.store, self.config, actor={'username':'operator'})
        self.assertEqual(self.sheet['sheet_period_end'], '2026-09-10')
        self.assertIn('2026-09-02', self.sheet['calendar_overrides'])


if __name__ == '__main__':
    unittest.main()
