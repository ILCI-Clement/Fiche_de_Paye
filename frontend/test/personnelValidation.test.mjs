import { test } from 'node:test'
import assert from 'node:assert/strict'
import { missingRegistrationFields, validateEmployment } from '../src/personnel/personnelValidation.ts'

const employee = { first_name: 'Synthetic', last_name: 'Fixture', nationality: 'Synthetic', sex: 'F',
  birth_date: '2000-01-01', address: 'Synthetic', job_title: 'Synthetic', qualification: 'Synthetic', personnel_type: 'salarie' }
const period = { entry_date: '2026-01-01', exit_date: null, actual_exit_date: null,
  contract_type: '', other_contract: '', is_foreign_worker: false, work_permit_type: '', work_permit_number: '',
  is_temporary_worker: false, temporary_company_name: '', temporary_company_address: '', is_young_worker: false, youth_contract_type: '' }

test('registration requires structured identity, basic fields and entry, not account or NIR', () => {
  assert.deepEqual(missingRegistrationFields(employee, period), [])
  for (const key of Object.keys(employee)) assert.ok(missingRegistrationFields({ ...employee, [key]: '  ' }, period).length)
  assert.deepEqual(missingRegistrationFields(employee, { ...period, entry_date: null }), ['Date d’entrée'])
})
test('employment validates actual and planned ends independently', () => {
  assert.equal(validateEmployment(period), '')
  for (const key of ['exit_date', 'actual_exit_date']) assert.ok(validateEmployment({ ...period, [key]: '2025-01-01' }))
})
test('all special categories enforce dependent fields and can coexist', () => {
  const complete = { ...period, is_foreign_worker: true, work_permit_type: 'Synthetic', work_permit_number: 'TEST',
    is_temporary_worker: true, contract_type: 'autre', other_contract: 'Synthetic', temporary_company_name: 'Synthetic',
    temporary_company_address: 'Synthetic', is_young_worker: true, youth_contract_type: 'apprenti' }
  assert.equal(validateEmployment(complete), '')
  for (const key of ['work_permit_type', 'work_permit_number', 'contract_type', 'other_contract',
    'temporary_company_name', 'temporary_company_address', 'youth_contract_type']) assert.ok(validateEmployment({ ...complete, [key]: '' }))
})
