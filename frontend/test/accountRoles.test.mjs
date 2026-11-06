import { test } from 'node:test'
import assert from 'node:assert/strict'
import { businessFunction, withBusinessFunction, withSystemAccess } from '../src/admin/accountRoles.ts'

test('business selection preserves system permissions and supports legacy mixed roles', () => {
  assert.equal(businessFunction(['Admin', 'Responsable', 'Employe']), 'Responsable')
  assert.deepEqual(withBusinessFunction(['SuperAdmin', 'Admin', 'Employe'], 'Responsable'), ['SuperAdmin', 'Admin', 'Responsable'])
  assert.deepEqual(withBusinessFunction(['Admin', 'Responsable'], 'Employe'), ['Admin', 'Employe'])
  assert.deepEqual(withBusinessFunction(['Admin', 'Employe'], ''), ['Admin'])
})

test('system checkboxes preserve business function and Super Admin inheritance', () => {
  assert.deepEqual(withSystemAccess(['Responsable'], 'Admin', true), ['Responsable', 'Admin'])
  assert.deepEqual(withSystemAccess(['Admin', 'Employe'], 'Admin', false), ['Employe'])
  assert.deepEqual(withSystemAccess(['Employe'], 'SuperAdmin', true), ['Employe', 'SuperAdmin', 'Admin'])
  assert.deepEqual(withSystemAccess(['SuperAdmin', 'Admin', 'Responsable'], 'SuperAdmin', false), ['Responsable', 'Admin'])
  assert.deepEqual(withSystemAccess(['SuperAdmin', 'Admin'], 'Admin', false), ['SuperAdmin', 'Admin'])
})
