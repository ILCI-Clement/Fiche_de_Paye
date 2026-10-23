"""New registration rules, legacy compatibility and date isolation on disposable data."""
from datetime import date
from io import BytesIO
import json
import os
import unittest
from openpyxl import load_workbook
from pydantic import ValidationError
from backend.personnel_rules import EmployeeCreate, PeriodInput, employment_status
from backend.personnel_contract_sync import sync_account_contracts
from test_support.personnel_fixtures import registration_payload
import test_personnel_register_sensitive as fixture
import test_personnel_contract_sync as contract_fixture
import test_personnel_sensitive_mysql as mysql_fixture


class RegistrationRulesTests(unittest.TestCase):
    def test_each_required_field_is_enforced_and_nir_account_exit_optional(self):
        valid = registration_payload()
        model = EmployeeCreate(**valid)
        self.assertIsNone(model.social_security_number)
        for key in ('first_name', 'last_name', 'nationality', 'sex', 'birth_date', 'address',
                    'job_title', 'qualification', 'personnel_type', 'initial_period'):
            with self.subTest(key=key), self.assertRaises(ValidationError):
                EmployeeCreate(**{**valid, key: None if key in ('birth_date', 'initial_period') else '  '})
        with self.assertRaises(ValidationError):
            EmployeeCreate(**{**valid, 'initial_period': {}})

    def test_special_categories_can_coexist_and_each_dependency_is_required(self):
        values = dict(entry_date='2026-01-01', contract_type='autre', other_contract='Synthetic',
            is_foreign_worker=True, work_permit_type='Synthetic permit', work_permit_number='TEST-ONLY',
            is_temporary_worker=True, temporary_company_name='Synthetic agency', temporary_company_address='Synthetic',
            is_young_worker=True, youth_contract_type='apprenti')
        self.assertTrue(PeriodInput(**values).is_foreign_worker)
        for key in ('contract_type', 'other_contract', 'work_permit_type', 'work_permit_number',
                    'temporary_company_name', 'temporary_company_address', 'youth_contract_type'):
            with self.subTest(key=key), self.assertRaises(ValidationError):
                PeriodInput(**{**values, key: ''})
        for flag in ('is_foreign_worker', 'is_temporary_worker', 'is_young_worker'):
            with self.subTest(flag=flag), self.assertRaises(ValidationError):
                PeriodInput(**{flag: 'false'})

    def test_departure_status_uses_actual_date_not_planned_end(self):
        period = dict(entry_date='2020-01-01', exit_date='2020-12-31')
        self.assertEqual(employment_status(period, date(2026, 1, 1)), 'employed')
        self.assertEqual(employment_status({**period, 'actual_exit_date': '2025-12-31'}, date(2026, 1, 1)), 'left')
        for key in ('actual_exit_date', 'exit_date'):
            with self.subTest(key=key), self.assertRaises(ValidationError):
                PeriodInput(entry_date='2026-01-01', **{key: '2025-01-01'})

    def test_model_errors_mask_nir_for_missing_fields_and_nested_errors(self):
        for change in ({'full_name': '', 'first_name': '', 'last_name': ''}, {'nationality': ''},
                       {'initial_period': {'is_foreign_worker': True}}):
            with self.subTest(change=change), self.assertRaises(ValidationError) as error:
                EmployeeCreate(**registration_payload(social_security_number=fixture.SYNTHETIC_NIR, **change))
            self.assertNotIn(fixture.SYNTHETIC_NIR, error.exception.json() + repr(error.exception.errors()))


class RegistrationHTTPTests(unittest.TestCase):
    setUp = fixture.PersonnelSensitiveHTTPTests.setUp
    create = fixture.PersonnelSensitiveHTTPTests.create
    update = fixture.PersonnelSensitiveHTTPTests.update
    rows = fixture.PersonnelSensitiveHTTPTests.rows
    raw_update = fixture.PersonnelSensitiveHTTPTests.raw_update

    def detail(self):
        result = self.create()
        self.assertEqual(result.status_code, 201, result.text)
        return result.json()

    def period_update(self, detail, **changes):
        employee, period = detail['employee'], detail['periods'][0]
        return self.client.put(f"/personnel-register/{employee['employee_id']}/periods/{period['period_id']}",
            json={'expected_version': employee['version'], **changes})

    def test_new_record_cannot_clear_mandatory_fields_or_spoof_schema(self):
        detail = self.detail()
        self.assertEqual(detail['employee']['registration_schema_version'], 2)
        for key in ('first_name', 'nationality', 'qualification'):
            self.assertEqual(self.update(detail['employee'], **{key: ''}).status_code, 400)
        self.assertEqual(self.period_update(detail, entry_date=None).status_code, 400)
        result = self.client.post('/personnel-register', json=registration_payload(registration_schema_version=1))
        self.assertEqual(result.status_code, 422)
        self.assertEqual(len(self.rows('personnel_register')), 1)

    def test_legacy_record_remains_editable_without_splitting_name(self):
        self.raw_update('INSERT INTO personnel_register(full_name) VALUES(?)', ('Unsplit historical name',))
        employee = self.client.get('/personnel-register/1').json()['employee']
        self.assertEqual(employee['first_name'], '')
        response = self.update(employee, job_title='Corrected job')
        self.assertEqual(response.status_code, 200, response.text)
        self.assertEqual(response.json()['employee']['full_name'], 'Unsplit historical name')
        self.assertTrue(response.json()['employee']['missing_fields'])

    def test_omitted_new_period_fields_survive_older_client_edits(self):
        detail = self.detail()
        result = self.period_update(detail, is_foreign_worker=True, work_permit_type='Synthetic', work_permit_number='TEST-ONLY')
        self.assertEqual(result.status_code, 200, result.text)
        result = self.period_update(result.json(), exit_date='2027-01-01')
        self.assertEqual(result.status_code, 200, result.text)
        self.assertEqual(result.json()['periods'][0]['work_permit_number'], 'TEST-ONLY')
        result = self.period_update(result.json(), is_foreign_worker=False)
        self.assertEqual(result.status_code, 200, result.text)
        self.assertEqual(result.json()['periods'][0]['work_permit_number'], '')
        self.assertIn('TEST-ONLY', json.dumps(self.rows('personnel_register_events')))

    def test_birth_and_entry_cannot_be_made_inconsistent_by_either_editor(self):
        detail = self.detail()
        self.assertEqual(self.period_update(detail, entry_date='1999-01-01').status_code, 400)
        self.assertEqual(self.update(detail['employee'], birth_date='2026-02-01').status_code, 400)
        response = self.client.post('/personnel-register/1/periods', json={
            'expected_version': detail['employee']['version'], 'entry_date': '1999-01-01'})
        self.assertEqual(response.status_code, 400)
        self.assertEqual(len(self.rows('personnel_register_periods')), 1)

    def test_period_writes_recheck_database_role_and_ownership(self):
        detail = self.detail()
        second = self.detail()
        foreign_period = second['periods'][0]['period_id']
        response = self.client.put(f"/personnel-register/1/periods/{foreign_period}", json={
            'expected_version': detail['employee']['version'], 'actual_exit_date': '2026-10-01'})
        self.assertEqual(response.status_code, 404)
        baseline = self.rows('personnel_register_periods')
        for name, role in (('fixture-manager', 'Responsable'), ('fixture-employee', 'Employe')):
            # A forged stale Admin claim cannot override the role in the database.
            self.actor = {'username': name, 'role_tags': ['Admin'], 'session_version': 0}
            response = self.period_update(detail, is_young_worker=True, youth_contract_type='apprenti')
            self.assertEqual(response.status_code, 403, response.text)
            response = self.client.post('/personnel-register/1/periods', json={
                'expected_version': detail['employee']['version'], 'entry_date': '2028-01-01'})
            self.assertEqual(response.status_code, 403)
        self.assertEqual(self.rows('personnel_register_periods'), baseline)

    def test_invalid_combined_period_patch_preserves_data_and_audit(self):
        detail = self.detail()
        response = self.period_update(detail, actual_exit_date='2026-06-01')
        self.assertEqual(response.status_code, 200, response.text)
        before = self.rows('personnel_register_periods')
        events = self.rows('personnel_register_events')
        response = self.period_update(response.json(), entry_date='2026-07-01')
        self.assertEqual(response.status_code, 400, response.text)
        self.assertEqual(self.rows('personnel_register_periods'), before)
        self.assertEqual(self.rows('personnel_register_events'), events)

    def test_special_fields_and_actual_departure_export_to_template(self):
        result = self.create(initial_period=dict(entry_date='2026-01-01', exit_date='2027-01-01',
            actual_exit_date='2026-10-01', contract_type='CDD', is_foreign_worker=True,
            work_permit_type='Synthetic permit', work_permit_number='TEST-ONLY', is_temporary_worker=True,
            temporary_company_name='Synthetic agency', temporary_company_address='Synthetic address',
            is_young_worker=True, youth_contract_type='apprenti'))
        self.assertEqual(result.status_code, 201, result.text)
        response = self.client.post('/personnel-register/export.xlsx', json={'confirm_sensitive_export': True})
        self.assertEqual(response.status_code, 200, response.text[:100])
        book = load_workbook(BytesIO(response.content))
        self.addCleanup(book.close)
        row = book.active
        self.assertEqual(row['K8'].value.strftime('%Y-%m-%d'), '2026-10-01')
        self.assertEqual(row['L8'].value, 'Synthetic permit')
        self.assertEqual(row['M8'].value, 'TEST-ONLY')
        self.assertEqual(row['N8'].value, 'CDD')
        self.assertIn('Synthetic agency', row['O8'].value)
        self.assertEqual(row['P8'].value, 'Oui')


class RegistrationSyncTests(unittest.TestCase):
    setUp = contract_fixture.ContractDateSyncTests.setUp
    create = fixture.PersonnelSensitiveHTTPTests.create
    rows = fixture.PersonnelSensitiveHTTPTests.rows
    raw_update = fixture.PersonnelSensitiveHTTPTests.raw_update
    seed_contract = contract_fixture.ContractDateSyncTests.seed_contract
    period_update = RegistrationHTTPTests.period_update

    def linked(self):
        self.seed_contract()
        result = self.create(linked_username='fixture-employee')
        self.assertEqual(result.status_code, 201, result.text)
        return result.json()

    def test_actual_departure_does_not_change_account_contract_even_when_archived(self):
        detail = self.linked()
        before = self.rows('personnel_contracts')
        self.raw_update('UPDATE personnel_contracts SET is_archived=1', ())
        result = self.period_update(detail, actual_exit_date='2026-10-01')
        self.assertEqual(result.status_code, 200, result.text)
        self.assertEqual(self.rows('personnel_contracts')[0]['end_date'], before[0]['end_date'])
        self.assertEqual(result.json()['periods'][0]['actual_exit_date'], '2026-10-01')

    def test_account_date_change_cannot_cross_actual_departure_and_rolls_back(self):
        detail = self.linked()
        result = self.period_update(detail, actual_exit_date='2026-10-01')
        self.assertEqual(result.status_code, 200, result.text)
        before = self.rows('personnel_contracts')
        response = self.client.patch(f"/users/fixture-employee/contracts/{before[0]['id']}", json={
            'employee_type': 'salarie', 'start_date': '2026-11-01', 'end_date': '2027-12-31', 'is_cdi': False})
        self.assertEqual(response.status_code, 409, response.text)
        self.assertEqual(self.rows('personnel_contracts'), before)

    def test_account_date_change_cannot_cross_birth_and_rolls_back(self):
        self.linked()
        before = self.rows('personnel_contracts')
        response = self.client.patch(f"/users/fixture-employee/contracts/{before[0]['id']}", json={
            'employee_type': 'salarie', 'start_date': '1999-01-01', 'end_date': '2027-12-31', 'is_cdi': False})
        self.assertEqual(response.status_code, 409, response.text)
        self.assertEqual(self.rows('personnel_contracts'), before)

    def test_startup_reconciliation_does_not_guess_new_unmatched_period(self):
        detail = self.linked()
        response = self.client.post(f"/personnel-register/{detail['employee']['employee_id']}/periods", json={
            'expected_version': detail['employee']['version'], 'entry_date': '2028-01-01'})
        self.assertEqual(response.status_code, 201, response.text)
        old = self.rows('personnel_register_periods')
        connection = fixture.Connection(self)
        try:
            with connection.cursor() as cursor:
                sync_account_contracts(cursor, 'fixture-employee', 'fixture-admin')
            connection.commit()
        finally:
            connection.close()
        self.assertEqual(self.rows('personnel_register_periods'), old)


@unittest.skipUnless(os.getenv('PRESENCE_MYSQL_TEMP_TEST') == '1', 'opt-in local MySQL temporary-table integration')
class RegistrationMySQLTests(unittest.TestCase):
    setUp = mysql_fixture.PersonnelSensitiveMySQLTests.setUp
    request = mysql_fixture.PersonnelSensitiveMySQLTests.request

    def test_special_fields_round_trip_and_actual_departure_filter(self):
        response = self.request('POST', '', json=registration_payload(initial_period={
            'entry_date': '2026-01-01', 'exit_date': '2027-12-31', 'actual_exit_date': '2026-02-01',
            'is_foreign_worker': True, 'work_permit_type': 'Synthetic permit', 'work_permit_number': 'TEST-ONLY',
            'is_temporary_worker': True, 'contract_type': 'CDD', 'temporary_company_name': 'Synthetic',
            'temporary_company_address': 'Synthetic address', 'is_young_worker': True, 'youth_contract_type': 'apprenti'}))
        self.assertEqual(response.status_code, 201, response.text)
        detail = response.json()
        identity = detail['employee']['employee_id']
        period = detail['periods'][0]
        self.assertIs(period['is_foreign_worker'], True)
        self.assertEqual(period['work_permit_number'], 'TEST-ONLY')
        self.assertEqual(self.request('GET', '?status=left').json()['total'], 1)
        self.assertEqual(self.request('GET', '?status=employed').json()['total'], 0)
        response = self.request('PUT', f"/{identity}/periods/{period['period_id']}", json={
            'expected_version': detail['employee']['version'], 'actual_exit_date': None})
        self.assertEqual(response.status_code, 200, response.text)
        self.assertIs(response.json()['periods'][0]['is_temporary_worker'], True)
        self.assertEqual(self.request('GET', '?status=employed').json()['total'], 1)

    def test_additive_schema_defaults_do_not_invent_structured_names_or_departure(self):
        with self.proxy.cursor() as cursor:
            cursor.execute('INSERT INTO personnel_register(full_name) VALUES (%s)', ('Unsplit synthetic history',))
            identity = cursor.lastrowid
            cursor.execute('INSERT INTO personnel_register_periods(employee_id,entry_date,exit_date) VALUES (%s,%s,%s)',
                           (identity, '2020-01-01', '2020-12-31'))
        self.proxy.commit()
        result = self.request('GET', f'/{identity}')
        self.assertEqual(result.status_code, 200, result.text)
        self.assertEqual(result.json()['employee']['registration_schema_version'], 1)
        self.assertEqual(result.json()['employee']['first_name'], '')
        self.assertIsNone(result.json()['periods'][0]['actual_exit_date'])
