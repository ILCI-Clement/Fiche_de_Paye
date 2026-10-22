import { test } from 'node:test'
import assert from 'node:assert/strict'
import { accessRoleLabel, canManageTeam, displayFunctionRoles, displayRoles, isAdmin, isSuperAdmin, roleLabel } from '../src/rolePolicy.ts'

test('administrative display labels use the agreed compact wording', () => {
  assert.equal(roleLabel('Admin'), 'Admin')
  assert.equal(roleLabel('SuperAdmin'), 'Super Admin')
  assert.equal(roleLabel('Responsable'), 'Responsable')
  assert.equal(roleLabel('Employe'), 'Employé')
})

test('role display has a stable order and hides inherited Admin under Super Admin', () => {
  assert.equal(displayRoles(['Employe', 'Admin', 'SuperAdmin', 'SuperAdmin']), 'Super Admin · Employé')
  assert.equal(displayRoles(['Employe', 'Responsable', 'Admin']), 'Admin · Responsable · Employé')
})

test('access and business function columns never merge their role categories', () => {
  const roles = ['Employe', 'Responsable', 'Admin', 'SuperAdmin']
  assert.equal(accessRoleLabel(roles), 'Super Admin')
  assert.equal(displayFunctionRoles(roles), 'Responsable · Employé')
  assert.equal(accessRoleLabel(['Employe']), 'Standard')
  assert.equal(displayFunctionRoles(['Admin']), '')
})

test('short display labels do not change internal authorization identifiers', () => {
  assert.equal(isSuperAdmin({ role_tags: ['SuperAdmin'] }), true)
  assert.equal(isAdmin({ role_tags: ['SuperAdmin'] }), true)
  assert.equal(isAdmin({ role_tags: ['Admin'] }), true)
  assert.equal(canManageTeam({ role_tags: ['Responsable'] }), true)
  assert.equal(isAdmin({ role_tags: ['Employe'] }), false)
  assert.equal(isSuperAdmin({ role_tags: ['Super Admin'] }), false)
})
