"""Account/register association regression using disposable SQLite fixtures only."""
import json
import unittest

import test_personnel_register_sensitive as fixture


class AccountPersonnelHTTPTests(unittest.TestCase):
    setUp = fixture.PersonnelSensitiveHTTPTests.setUp
    create = fixture.PersonnelSensitiveHTTPTests.create
    rows = fixture.PersonnelSensitiveHTTPTests.rows
    stored = fixture.PersonnelSensitiveHTTPTests.stored
    raw_update = fixture.PersonnelSensitiveHTTPTests.raw_update
    assert_safe = fixture.PersonnelSensitiveHTTPTests.assert_safe
    assert_no_cache = fixture.PersonnelSensitiveHTTPTests.assert_no_cache

    target = 'fixture-employee'

    def employee(self, **values):
        response = self.create(**values)
        self.assertEqual(response.status_code, 201, response.text)
        return response.json()['employee']

    def link(self, employee=None, current=None, **overrides):
        body = {
            'employee_id': employee['employee_id'] if employee else None,
            'target_version': employee['version'] if employee else None,
            'expected_employee_id': current['employee_id'] if current else None,
            'expected_version': current['version'] if current else None,
            **overrides,
        }
        return self.client.put(f'/account-personnel/{self.target}/link', json=body)

    def get(self):
        return self.client.get(f'/account-personnel/{self.target}')

    def test_associate_read_and_unlink_preserve_records_and_history(self):
        employee = self.employee(initial_period={'entry_date': '2026-01-01', 'contract_type': 'CDI'})
        result = self.link(employee)
        self.assertEqual(result.status_code, 200, result.text)
        self.assert_no_cache(result)
        linked = result.json()['employee']
        self.assertEqual(linked['linked_username'], self.target)
        self.assertEqual(linked['version'], employee['version'] + 1)
        self.assertEqual(len(result.json()['periods']), 1)
        result = self.link(current=linked)
        self.assertEqual(result.status_code, 200, result.text)
        self.assertIsNone(result.json()['employee'])
        self.assertEqual(len(self.rows('personnel_register')), 1)
        self.assertEqual(len(self.rows('personnel_register_periods')), 1)
        self.assertEqual([row['action'] for row in self.rows('personnel_register_events')],
                         ['registered', 'account_linked', 'account_unlinked'])

    def test_self_is_read_only_and_does_not_receive_sensitive_fields(self):
        employee = self.employee(social_security_number=fixture.SYNTHETIC_NIR)
        self.assertEqual(self.link(employee).status_code, 200)
        self.actor = {'username': self.target, 'role_tags': ['Employe'], 'session_version': 0}
        result = self.get()
        self.assertEqual(result.status_code, 200, result.text)
        self.assertFalse(result.json()['can_manage_link'])
        self.assert_safe(result.json())
        self.assertNotIn('has_social_security_number', result.text)
        self.assert_no_cache(result)
        self.assertEqual(self.link(current=result.json()['employee']).status_code, 403)
        self.assertEqual(self.client.get(f'/account-personnel/{self.target}/choices').status_code, 403)
        self.assertEqual(self.client.get('/personnel-register').status_code, 403)
        self.assertEqual(self.client.get('/account-personnel/fixture-admin').status_code, 403)

    def test_manager_cannot_read_another_persons_private_register_projection(self):
        self.actor = {'username': 'fixture-manager', 'role_tags': ['Responsable'], 'session_version': 0}
        self.assertEqual(self.get().status_code, 403)
        self.assertEqual(self.link().status_code, 403)

    def test_replacement_is_atomic_and_audits_both_records(self):
        first, second = self.employee(), self.employee()
        first = self.link(first).json()['employee']
        result = self.link(second, first)
        self.assertEqual(result.status_code, 200, result.text)
        self.assertIsNone(self.stored(first)['linked_username'])
        self.assertEqual(self.stored(second)['linked_username'], self.target)
        events = self.rows('personnel_register_events')[-2:]
        self.assertEqual([event['action'] for event in events], ['account_unlinked', 'account_linked'])
        self.assertTrue(all(event['actor_username'] == 'fixture-admin' for event in events))

    def test_audit_failure_rolls_back_both_sides_of_replacement(self):
        first, second = self.employee(), self.employee()
        first = self.link(first).json()['employee']
        baseline = self.rows('personnel_register')
        events = self.rows('personnel_register_events')
        self.fail_audit = True
        with self.assertLogs('backend.personnel_register', level='ERROR'):
            result = self.link(second, first)
        self.assertEqual(result.status_code, 500, result.text)
        self.assertEqual(self.rows('personnel_register'), baseline)
        self.assertEqual(self.rows('personnel_register_events'), events)

    def test_stale_current_and_target_versions_do_not_mutate(self):
        first, second = self.employee(), self.employee()
        first = self.link(first).json()['employee']
        baseline = self.rows('personnel_register')
        for overrides in ({'expected_employee_id': None, 'expected_version': None},
                          {'expected_version': 1}, {'target_version': 99}):
            self.assertEqual(self.link(second, first, **overrides).status_code, 409)
        self.assertEqual(self.rows('personnel_register'), baseline)

    def test_occupied_and_archived_targets_are_rejected(self):
        occupied = self.employee(linked_username='fixture-manager')
        archived = self.employee()
        self.raw_update('UPDATE personnel_register SET is_archived=1 WHERE employee_id=?',
                        (archived['employee_id'],))
        for employee in (occupied, archived):
            self.assertEqual(self.link(employee).status_code, 409)
        self.assertIsNone(self.get().json()['employee'])

    def test_choices_only_offer_unarchived_unoccupied_records(self):
        available = self.employee(full_name='Available')
        self.employee(full_name='Occupied', linked_username='fixture-manager')
        archived = self.employee(full_name='Archived')
        self.raw_update('UPDATE personnel_register SET is_archived=1 WHERE employee_id=?',
                        (archived['employee_id'],))
        result = self.client.get(f'/account-personnel/{self.target}/choices')
        self.assertEqual(result.status_code, 200, result.text)
        self.assertEqual([item['employee_id'] for item in result.json()['employees']], [available['employee_id']])
        self.assert_no_cache(result)
        self.assertNotIn('address', result.text)
        self.assert_safe(result.json())
        result = self.client.get(f'/account-personnel/{self.target}/choices', params={'q': "%' OR 1=1 --"})
        self.assertEqual(result.status_code, 200)
        self.assertEqual(result.json()['employees'], [])

    def test_fresh_database_authority_and_session_version_are_required(self):
        employee = self.employee()
        for changes in ({'role_tags': json.dumps(['Employe'])}, {'is_active': 0},
                        {'session_version': 1}, {'login_deleted_at': '2026-01-01'}):
            field, value = next(iter(changes.items()))
            previous = next(row[field] for row in self.rows('users') if row['username'] == 'fixture-admin')
            self.raw_update(f'UPDATE users SET {field}=? WHERE username=?', (value, 'fixture-admin'))
            self.assertEqual(self.link(employee).status_code, 403 if field == 'role_tags' else 401)
            self.raw_update(f'UPDATE users SET {field}=? WHERE username=?', (previous, 'fixture-admin'))
        self.assertIsNone(self.stored(employee)['linked_username'])

    def test_same_association_is_noop_and_invalid_payload_is_redacted(self):
        employee = self.link(self.employee()).json()['employee']
        events = self.rows('personnel_register_events')
        result = self.link(employee, employee)
        self.assertEqual(result.status_code, 200)
        self.assertEqual(self.rows('personnel_register_events'), events)
        result = self.link(employee, employee, employee_id='private-invalid-identifier')
        self.assertEqual(result.status_code, 422)
        self.assertNotIn('private-invalid-identifier', result.text)
        self.assert_no_cache(result)


if __name__ == '__main__':
    unittest.main()
