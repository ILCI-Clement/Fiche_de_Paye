import { test } from 'node:test'
import assert from 'node:assert/strict'
import { selectedAccount } from '../src/admin/accountSelection.ts'

const accounts = [
  { username: 'christophe', display_name: 'Christophe' },
  { username: 'emilie-key', display_name: 'Emilie' },
  { username: 'another-emilie', display_name: 'Emilie' },
]

test('dashboard targets exact account identity even when names duplicate or order changes', () => {
  assert.equal(selectedAccount(accounts, 'emilie-key'), accounts[1])
  assert.equal(selectedAccount([...accounts].reverse(), 'emilie-key'), accounts[1])
})

test('inaccessible or removed navigation targets never open the first account', () => {
  assert.equal(selectedAccount(accounts, 'hidden-or-deleted'), undefined)
  assert.equal(selectedAccount([], 'emilie-key'), undefined)
  assert.equal(selectedAccount(accounts, 'EMILIE-KEY'), undefined)
  assert.equal(selectedAccount(accounts, ''), accounts[0])
})
