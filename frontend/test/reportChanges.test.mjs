import { test } from 'node:test'
import assert from 'node:assert/strict'
import { reportChanges } from '../src/admin/reportChanges.ts'

const accounts = [{username:'current',manager_id:'boss'}, {username:'transfer',manager_id:'other'}, {username:'new',manager_id:null}]
test('batch edits retain prior manager for conflict checks and permit no manager', () => {
  assert.deepEqual(reportChanges(accounts,'boss',['transfer','new']), [
    {username:'current',expected_manager_id:'boss',assigned:false},
    {username:'transfer',expected_manager_id:'other',assigned:true},
    {username:'new',expected_manager_id:null,assigned:true},
  ])
})
test('unchanged selection sends nothing and unavailable accounts are never removed', () => {
  assert.deepEqual(reportChanges(accounts,'boss',['current']), [])
  assert.deepEqual(reportChanges(accounts.slice(1),'boss',[]), [])
})
