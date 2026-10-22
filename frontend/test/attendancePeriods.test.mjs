import { test } from 'node:test'
import assert from 'node:assert/strict'
import { accountSheetPeriods } from '../src/attendance/types.ts'

test('existing month navigation uses explicit account identity and preserves records', () => {
  const fiches = [
    { id: 'sep', account_username: 'id-a', sheet_month: 9, sheet_year: 2026, nom: 'Shared Name' },
    { id: 'other', account_username: 'id-b', sheet_month: 10, sheet_year: 2026, nom: 'Shared Name' },
    { id: 'unlinked', email_employe: 'a@example.test', sheet_month: 8, sheet_year: 2026 },
  ]
  const before = structuredClone(fiches)
  assert.deepEqual(accountSheetPeriods(fiches, 'id-a'), [{ month: 9, year: 2026, id: 'sep' }])
  assert.deepEqual(fiches, before)
  assert.deepEqual(accountSheetPeriods(fiches, ''), [])
})

test('existing periods are deduplicated and sorted newest first across years', () => {
  const fiches = [
    { id: 'first-sep', account_username: 'id-a', sheet_month: 9, sheet_year: 2026 },
    { id: 'second-sep-contract', account_username: 'id-a', sheet_month: 9, sheet_year: 2026 },
    { id: 'new-year', account_username: 'id-a', sheet_month: 1, sheet_year: 2027 },
    { id: 'oct', account_username: 'id-a', sheet_month: 10, sheet_year: 2026 },
  ]
  assert.deepEqual(accountSheetPeriods(fiches, 'id-a').map(item => item.id), ['new-year', 'oct', 'first-sep'])
})

test('archived and normal months remain separate', () => {
  const fiches = [
    { id: 'normal', account_username: 'id-a', sheet_month: 9, sheet_year: 2026 },
    { id: 'archived', account_username: 'id-a', sheet_month: 8, sheet_year: 2026, is_archived: true },
  ]
  assert.deepEqual(accountSheetPeriods(fiches, 'id-a').map(item => item.id), ['normal'])
  assert.deepEqual(accountSheetPeriods(fiches, 'id-a', true).map(item => item.id), ['archived'])
})

test('legacy unbound and invalid dates do not invent saved months', () => {
  const fiches = [{ id: 'legacy', account_username: 'id-a' },
    ...[0, 13, '9', null, undefined, 9.5].map((month, i) => ({ id: `invalid-${i}`, account_username: 'id-a', sheet_month: month, sheet_year: 2026 })),
    { id: 'invalid-year', account_username: 'id-a', sheet_month: 9, sheet_year: '2026' }]
  assert.deepEqual(accountSheetPeriods(fiches, 'id-a'), [])
})
