import { test } from 'node:test'
import assert from 'node:assert/strict'
import { completionSection, missingPersonnelField } from '../src/completionTarget.ts'

test('completion routes preserve the actual issue domain, including manager contact identity', () => {
  assert.equal(completionSection('manager'), 'organization')
  assert.equal(completionSection('manager_contact'), 'register')
  assert.equal(completionSection('contract'), 'contract')
  assert.equal(completionSection('register_link'), 'register')
})

test('focus first missing basic field, or explicitly requested email, without selecting NIR', () => {
  const person = { last_name: 'Test', first_name: '', personnel_type: 'salarie' }
  assert.equal(missingPersonnelField(person), 'first_name')
  assert.equal(missingPersonnelField(person, 'email'), 'email')
  assert.notEqual(missingPersonnelField({}), 'social_security_number')
  assert.equal(missingPersonnelField({ last_name: 'Test', first_name: 'Test', personnel_type: 'salarie', nationality: 'Test', sex: 'F', birth_date: '2000-01-01', address: 'Test', job_title: 'Test', qualification: 'Test' }), undefined)
})
