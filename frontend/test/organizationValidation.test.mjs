import { test } from 'node:test'
import assert from 'node:assert/strict'
import { organizationValidation } from '../src/admin/organizationValidation.ts'

const value = { role_tags: ['Responsable', 'Employe'], employee_type: 'salarie', manager_id: 'boss', group_ids: [1], managed_group_ids: [1] }
test('a department membership does not replace the required managed department', () => {
  assert.match(organizationValidation({ ...value, managed_group_ids: [] }), /département géré/)
  assert.equal(organizationValidation(value), '')
})
test('employee function requires its own direct manager, including a manager who is also an employee', () => {
  assert.match(organizationValidation({ ...value, manager_id: null }), /responsable direct/)
  assert.equal(organizationValidation({ ...value, role_tags: ['Responsable'], manager_id: null }), '')
})
test('admin inherits general scope but department management still requires the business function', () => {
  assert.equal(organizationValidation({ ...value, role_tags: ['Admin', 'Responsable'], managed_group_ids: [], manager_id: null }), '')
  assert.equal(organizationValidation({ ...value, role_tags: ['SuperAdmin', 'Responsable'], managed_group_ids: [], manager_id: null }), '')
  assert.match(organizationValidation({ ...value, role_tags: ['Employe'] }), /fonction Responsable/)
})
