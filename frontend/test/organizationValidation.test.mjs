import { test } from 'node:test'
import assert from 'node:assert/strict'
import { organizationValidation } from '../src/admin/organizationValidation.ts'

const base = { role_tags: ['Employe'], employee_type: 'salarie', manager_id: null, group_ids: [], managed_group_ids: [] }
test('an employee account does not require a direct manager', () => {
  assert.equal(organizationValidation(base), '')
})
test('a system Super Admin does not require business bindings', () => {
  assert.equal(organizationValidation({ ...base, role_tags: ['SuperAdmin', 'Admin'] }), '')
  assert.equal(organizationValidation({ ...base, role_tags: ['SuperAdmin', 'Admin', 'Employe'] }), '')
})
test('ordinary Responsable department and role validation remain enforced', () => {
  assert.notEqual(organizationValidation({ ...base, role_tags: ['Responsable'] }), '')
  assert.notEqual(organizationValidation({ ...base, managed_group_ids: [1] }), '')
  assert.notEqual(organizationValidation({ ...base, role_tags: [] }), '')
})
