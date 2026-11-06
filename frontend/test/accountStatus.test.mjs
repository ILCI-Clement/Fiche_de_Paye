import { test } from 'node:test'
import assert from 'node:assert/strict'
import { accountStatus } from '../src/admin/accountStatus.ts'

test('legacy and normal accounts are active; disabled accounts stay separately recoverable', () => {
  assert.equal(accountStatus({}), 'active')
  assert.equal(accountStatus({ is_active: true }), 'active')
  assert.equal(accountStatus({ is_active: false }), 'inactive')
})
test('deleted login always belongs to history even if an old active flag is inconsistent', () => {
  assert.equal(accountStatus({ is_active: false, login_deleted_at: '2026-10-29' }), 'deleted')
  assert.equal(accountStatus({ is_active: true, login_deleted_at: '2026-10-29' }), 'deleted')
})

test('archived employees leave daily catalogs without changing login state', () => {
  assert.equal(accountStatus({ is_active: true, personnel_archived: true }), 'archived')
  assert.equal(accountStatus({ is_active: false, personnel_archived: true }), 'inactive')
  assert.equal(accountStatus({ personnel_archived: true, login_deleted_at: '2026-10-29' }), 'deleted')
})
