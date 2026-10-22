"""Contract/register HTTP regressions on disposable SQLite fixtures, no real users."""
from backend.contract_history import register_contract_history
from backend.personnel_contract_sync import sync_account_contracts
import unittest
import os
from datetime import date
from fastapi import HTTPException
from backend.personnel_contract_sync import sync_register_period
import test_personnel_sensitive_mysql as mysql_fixture
import test_personnel_register_sensitive as fixture


class ContractDateSyncTests(unittest.TestCase):
    create = fixture.PersonnelSensitiveHTTPTests.create
    rows = fixture.PersonnelSensitiveHTTPTests.rows
    raw_update = fixture.PersonnelSensitiveHTTPTests.raw_update
    def setUp(self):
        fixture.PersonnelSensitiveHTTPTests.setUp(self)
        def load_user(cursor, username):
            cursor.execute('SELECT * FROM users WHERE username=%s', (username,))
            return cursor.fetchone()
        register_contract_history(self.client.app, lambda: fixture.Connection(self),
                                  lambda *_: lambda: dict(self.actor), lambda: dict(self.actor),
                                  lambda *_args, **_kwargs: None, load_user=load_user)

    def seed_contract(self, start='2026-01-01', end='2027-12-31'):
        self.raw_update('INSERT INTO personnel_contracts(username,employee_type,start_date,end_date,is_cdi) VALUES(?,?,?,?,?)',
                        ('fixture-employee', 'salarie', start, end, int(end is None)))
        return self.rows('personnel_contracts')[-1]['id']

    def linked(self, **extra):
        response = self.create(linked_username='fixture-employee', **extra)
        self.assertEqual(response.status_code, 201, response.text)
        return response.json()

    def save_period(self, detail, start, end):
        period = detail['periods'][0]
        return self.client.put(f"/personnel-register/{detail['employee']['employee_id']}/periods/{period['period_id']}", json={
            'expected_version': detail['employee']['version'], 'entry_date': start, 'exit_date': end,
            'contract_type': period['contract_type'], 'other_contract': period['other_contract'],
        })

    def test_initial_account_dates_win_and_second_sync_is_idempotent(self):
        identity = self.seed_contract()
        detail = self.linked(initial_period={'entry_date': '2020-01-01', 'exit_date': '2021-01-01'})
        self.assertEqual(len(detail['periods']), 1)
        self.assertEqual(detail['periods'][0]['entry_date'], '2026-01-01')
        self.assertEqual(detail['periods'][0]['exit_date'], '2027-12-31')
        self.assertEqual(detail['periods'][0]['account_contract_id'], identity)
        count = len(self.rows('personnel_register_events'))
        connection = fixture.Connection(self)
        try:
            with connection.cursor() as cursor:
                sync_account_contracts(cursor, 'fixture-employee', 'fixture-admin')
            connection.commit()
        finally:
            connection.close()
        self.assertEqual(count, len(self.rows('personnel_register_events')))

    def test_register_to_account_and_account_to_register_preserve_ids_and_stale_guard(self):
        identity = self.seed_contract()
        detail = self.linked()
        response = self.save_period(detail, '2026-02-01', '2027-11-30')
        self.assertEqual(response.status_code, 200, response.text)
        self.assertEqual(self.rows('personnel_contracts')[0]['start_date'], '2026-02-01')
        self.assertEqual(self.rows('personnel_contracts')[0]['id'], identity)
        self.assertEqual(self.rows('users')[2]['contract_start_date'], '2026-02-01')
        account = self.client.patch(f'/users/fixture-employee/contracts/{identity}', json={
            'employee_type': 'salarie', 'start_date': '2026-03-01', 'end_date': '2027-10-31', 'is_cdi': False,
        })
        self.assertEqual(account.status_code, 200, account.text)
        latest = self.client.get(f"/personnel-register/{detail['employee']['employee_id']}").json()
        self.assertEqual(latest['periods'][0]['entry_date'], '2026-03-01')
        self.assertEqual(latest['periods'][0]['period_id'], detail['periods'][0]['period_id'])
        self.assertEqual(self.save_period(detail, '2026-04-01', '2027-09-30').status_code, 409)

    def test_overlapping_missing_end_and_archived_contract_edits_roll_back(self):
        self.seed_contract('2026-01-01', '2026-06-30')
        identity = self.seed_contract('2026-07-01', '2027-12-31')
        detail = self.linked()
        old = self.rows('personnel_contracts')
        self.assertEqual(self.save_period(detail, '2026-06-01', '2027-12-31').status_code, 409)
        self.assertEqual(self.save_period(detail, '2026-07-01', None).status_code, 400)
        self.assertEqual(self.rows('personnel_contracts'), old)
        self.raw_update('UPDATE personnel_contracts SET is_archived=1 WHERE id=?', (identity,))
        self.assertEqual(self.save_period(detail, '2026-07-02', '2027-12-31').status_code, 409)

    def test_audit_failure_rolls_back_both_directions(self):
        identity = self.seed_contract()
        detail = self.linked()
        old_contracts = self.rows('personnel_contracts')
        old_periods = self.rows('personnel_register_periods')
        self.fail_audit = True
        with self.assertLogs('backend.personnel_register', level='ERROR'):
            self.assertEqual(self.save_period(detail, '2026-02-01', '2027-11-30').status_code, 500)
        self.assertEqual(self.rows('personnel_contracts'), old_contracts)
        self.assertEqual(self.rows('personnel_register_periods'), old_periods)

        account = self.client.patch(f'/users/fixture-employee/contracts/{identity}', json={
            'employee_type': 'salarie', 'start_date': '2026-03-01', 'end_date': '2027-11-30', 'is_cdi': False,
        })
        self.assertEqual(account.status_code, 500)
        self.assertEqual(self.rows('personnel_contracts'), old_contracts)
        self.assertEqual(self.rows('personnel_register_periods'), old_periods)


@unittest.skipUnless(os.getenv('PRESENCE_MYSQL_TEMP_TEST') == '1', 'opt-in local MySQL temporary-table integration')
class ContractDateMySQLTests(unittest.TestCase):
    setUp = mysql_fixture.PersonnelSensitiveMySQLTests.setUp

    def test_real_mysql_mapping_dates_overlap_and_no_extra_history_on_repeat(self):
        with self.proxy.cursor() as cursor:
            cursor.execute('''INSERT INTO personnel_contracts(username,employee_type,start_date,end_date,is_cdi)
                VALUES (%s,%s,%s,%s,%s)''', ('fixture-employee', 'salarie', '2026-01-01', '2026-06-30', False))
            contract_id = cursor.lastrowid
        self.proxy.commit()
        response = self.client.post('/personnel-register', headers=self.headers,
                                    json={'full_name': 'Temporary synthetic sync', 'linked_username': 'fixture-employee'})
        self.assertEqual(response.status_code, 201, response.text)
        detail = response.json()
        period = detail['periods'][0]
        self.assertEqual(period['account_contract_id'], contract_id)
        with self.proxy.cursor() as cursor:
            cursor.execute('UPDATE personnel_register_periods SET exit_date=%s WHERE period_id=%s', ('2026-05-31', period['period_id']))
            sync_register_period(cursor, detail['employee']['employee_id'], period['period_id'])
            cursor.execute('SELECT end_date FROM personnel_contracts WHERE id=%s', (contract_id,))
            self.assertEqual(cursor.fetchone()['end_date'], date(2026, 5, 31))
            cursor.execute('UPDATE personnel_contracts SET end_date=%s WHERE id=%s', ('2026-04-30', contract_id))
            sync_account_contracts(cursor, 'fixture-employee', 'fixture-admin')
            cursor.execute('SELECT exit_date FROM personnel_register_periods WHERE period_id=%s', (period['period_id'],))
            self.assertEqual(cursor.fetchone()['exit_date'], date(2026, 4, 30))
            cursor.execute('SELECT COUNT(*) AS total FROM personnel_register_events')
            count = cursor.fetchone()['total']
            sync_account_contracts(cursor, 'fixture-employee', 'fixture-admin')
            cursor.execute('SELECT COUNT(*) AS total FROM personnel_register_events')
            self.assertEqual(cursor.fetchone()['total'], count)
        self.proxy.commit()
        with self.proxy.cursor() as cursor:
            cursor.execute('''INSERT INTO personnel_contracts(username,employee_type,start_date,end_date,is_cdi)
                VALUES (%s,%s,%s,%s,%s)''', ('fixture-employee', 'salarie', '2026-07-01', '2026-12-31', False))
        self.proxy.commit()
        with self.proxy.cursor() as cursor:
            cursor.execute('UPDATE personnel_register_periods SET exit_date=%s WHERE period_id=%s', ('2026-08-31', period['period_id']))
            with self.assertRaises(HTTPException) as caught:
                sync_register_period(cursor, detail['employee']['employee_id'], period['period_id'])
            self.assertEqual(caught.exception.status_code, 409)
        self.proxy.rollback()
