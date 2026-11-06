import { test } from 'node:test'
import assert from 'node:assert/strict'
import { teamPeople } from '../src/team/teamPeople.ts'

test('team table includes pure managers alongside employees', () => {
  const rows = [
    { username: 'report-manager', role_tags: ['Responsable'] },
    { username: 'employee', role_tags: ['Employe'] },
    { username: 'mixed', role_tags: ['Admin', 'Responsable'] },
    { username: 'admin-only', role_tags: ['Admin'] },
    { username: 'legacy-manager', role: 'Responsable' },
  ]
  assert.deepEqual(teamPeople(rows).map(row => row.username), ['report-manager', 'employee', 'mixed', 'legacy-manager'])
})

test('manager identity does not bypass inactive, archived or deleted filters', () => {
  const rows = [
    { username: 'inactive', role_tags: ['Responsable'], is_active: false },
    { username: 'archived', role_tags: ['Responsable'], personnel_archived: true },
    { username: 'deleted', role_tags: ['Responsable'], login_deleted_at: '2026-10-29' },
  ]
  assert.deepEqual(teamPeople(rows), [])
})
