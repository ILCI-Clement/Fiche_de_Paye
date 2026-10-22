import { test } from 'node:test'
import assert from 'node:assert/strict'
import { QueryClient } from '@tanstack/react-query'
import { personDisplayName, personOptionLabel } from '../src/personDisplayName.ts'
import { accountConfirmation } from '../src/admin/accountIdentity.ts'
import { attendancePersonName, ficheLabel, withAttendanceLabels } from '../src/attendance/types.ts'
import { matchesDocumentSearch } from '../src/documents/documentSearch.ts'

const account = (username, email, full_name) => ({ username, email, full_name, contracts: [] })
const first = account('internal-one', 'one@example.invalid', 'Alex Martin')
const second = account('internal-two', 'two@example.invalid', 'Alex Martin')
const workspace = () => ({
  config: { mois: 10, annee: 2026, employes_data: [{ id: 'sheet-1', account_username: first.username, nom: 'Historical document name', email_employe: first.email }] },
  accounts: [first, second], revision: 'saved-revision', can_edit_identity: true, can_delete_sheets: false, can_archive_sheets: true,
  person_labels: { 'sheet-1': 'Alex Martin' },
})

test('registered full name has priority over projection and email', () => {
  assert.equal(personDisplayName({ full_name: '  Élodie Durand  ', display_name: 'Previous label', email: 'email-prefix@example.invalid' }), 'Élodie Durand')
})
test('a canonical projection is accepted when full name is absent', () => {
  assert.equal(personDisplayName({ full_name: '  ', display_name: ' Alex Martin ', email: first.email }), 'Alex Martin')
})
test('fallback is the full email local part with punctuation preserved', () => {
  assert.equal(personDisplayName({ email: ' christophe.huang+test@example.invalid ' }), 'christophe.huang+test')
})
for (const email of ['', '   ', 'no-domain', '@example.invalid', 'local@', 'local@@example.invalid', 'bad local@example.invalid', 'local@bad domain']) {
  test(`invalid email does not become a person label: ${JSON.stringify(email)}`, () => {
    assert.equal(personDisplayName({ username: 'private-login', name: 'legacy-username', email }), 'Nom non renseigné')
  })
}
test('missing person and raw legacy name never expose the username', () => {
  assert.equal(personDisplayName(null), 'Nom non renseigné')
  assert.equal(personDisplayName({ username: 'private-login', name: 'private-login' }), 'Nom non renseigné')
})
test('namesakes in selectors are distinguished by email without login identifiers', () => {
  assert.notEqual(personOptionLabel(first), personOptionLabel(second))
  assert.equal(personOptionLabel(first), 'Alex Martin · one@example.invalid')
  assert.ok(!personOptionLabel(first).includes(first.username))
})
test('selectors do not expose a username when email is missing', () => {
  assert.equal(personOptionLabel({ username: 'private-login', full_name: 'Alex Martin' }), 'Alex Martin')
})
test('destructive confirmation uses current email independently of display label', () => {
  assert.deepEqual(accountConfirmation({ ...first, email: ' one@example.invalid ' }), { expected: first.email, label: 'Adresse e-mail de confirmation' })
  assert.notEqual(accountConfirmation(first).expected, personDisplayName(first))
})
test('namesakes have distinct destructive confirmations', () => {
  assert.notEqual(accountConfirmation(first).expected, accountConfirmation(second).expected)
})
test('missing email uses an explicitly labeled technical confirmation, never the name', () => {
  assert.deepEqual(accountConfirmation({ ...first, email: ' ' }), { expected: first.username, label: 'Identifiant technique de confirmation' })
})
test('attendance resolves identity by account key, never by a duplicate name', () => {
  const data = workspace()
  const modified = { ...second, full_name: 'Second person renamed' }
  data.accounts[1] = modified
  assert.equal(attendancePersonName({ id: 'sheet-2', account_username: second.username }, data), modified.full_name)
  assert.equal(attendancePersonName(data.config.employes_data[0], data), first.full_name)
})
test('attendance display uses saved sheet metadata when a scoped catalog lacks its account', () => {
  const data = workspace()
  assert.equal(attendancePersonName(data.config.employes_data[0], { accounts: [], person_labels: data.person_labels }), first.full_name)
})
test('unlinked legacy attendance uses email prefix without historical username', () => {
  assert.equal(attendancePersonName({ id: 'legacy', nom: 'private-login', email_employe: 'prefix@example.invalid' }, { accounts: [] }), 'prefix')
  assert.equal(attendancePersonName({ id: 'legacy', nom: 'private-login' }, { accounts: [] }), 'Nom non renseigné')
})
test('label refresh keeps revision, contract snapshot and unsaved content unchanged', () => {
  const current = workspace()
  current.config.employes_data[0].calendar_overrides = { '2026-10-22': { matin: ['CP', ''] } }
  const before = structuredClone(current)
  const next = withAttendanceLabels(current, { person_labels: { 'sheet-1': 'Updated employee name' } })
  assert.equal(next.config, current.config)
  assert.equal(next.accounts, current.accounts)
  assert.equal(next.revision, current.revision)
  assert.deepEqual(current, before)
  assert.equal(attendancePersonName(next.config.employes_data[0], next), 'Updated employee name')
  assert.equal(next.config.employes_data[0].nom, 'Historical document name')
})
test('save responses without labels leave the metadata intact', () => {
  const current = workspace()
  assert.equal(withAttendanceLabels(current, {}), current)
})
test('sheet labels retain date suffixes and stable IDs', () => {
  const fiche = { id: 123, account_username: first.username, contract_id: 42, sheet_period_start: '2026-10-01', sheet_period_end: '2026-10-15', nom: 'Old export text' }
  assert.equal(ficheLabel(fiche, 'Alex Martin'), 'Alex Martin · 01–15')
  assert.equal(fiche.id, 123)
  assert.equal(fiche.nom, 'Old export text')
})
test('query invalidation caused by a name change leaves existing cache/draft content in place', async () => {
  const client = new QueryClient()
  const current = workspace()
  const key = ['attendance-workspace', first.username]
  client.setQueryData(key, current)
  client.setQueryData(['documents', first.username], { receipts: [] })
  await client.invalidateQueries()
  assert.equal(client.getQueryData(key), current)
  assert.equal(client.getQueryState(key).isInvalidated, true)
  assert.equal(client.getQueryState(['documents', first.username]).isInvalidated, true)
  client.clear()
})

const receipt = { id: 1, document_number: '20261022-0001', employee_username: 'private-login', uploaded_by: 'private-uploader', original_filename: 'Justificatif.pdf', employee_display_name: 'Élodie Durand', uploaded_by_display_name: 'Alex Martin', archived_by_display_name: 'Sam Dupont', document_type: 'autre', other_type: 'Hôtel Paris', size_bytes: 1024, receipt_month: 10, receipt_year: 2026, archived: false, created_at: '2026-10-22T12:00:00Z' }
test('document search matches canonical employee and uploader names without case or accents', () => {
  assert.equal(matchesDocumentSearch(receipt, 'ELODIE durand alex'), true)
  assert.equal(matchesDocumentSearch(receipt, 'dupont'), true)
})
test('document search uses numbers, filename and category as before', () => {
  assert.equal(matchesDocumentSearch(receipt, '20261022-0001 justificatif hotel'), true)
  assert.equal(matchesDocumentSearch(receipt, 'absent'), false)
})
test('document search excludes technical employee and actor identifiers', () => {
  assert.equal(matchesDocumentSearch(receipt, 'private-login'), false)
  assert.equal(matchesDocumentSearch(receipt, 'private-uploader'), false)
})
test('document search handles missing projections and empty search without exposing IDs', () => {
  assert.equal(matchesDocumentSearch({ ...receipt, employee_display_name: undefined, uploaded_by_display_name: undefined }, ' '), true)
  assert.equal(matchesDocumentSearch({ ...receipt, employee_display_name: undefined }, 'private-login'), false)
})
