import { test } from 'node:test'
import assert from 'node:assert/strict'
import { canAdoptWorkspaceRevision, refreshWorkspaceContracts } from '../src/attendance/contractRefresh.ts'

test('current and protected workspaces do not trigger a write', async () => {
  const workspace = { config: { employes_data: [] }, revision: 'current' }
  assert.equal(await refreshWorkspaceContracts(workspace, () => { throw Error('Unexpected write') }), workspace)
})

test('stale snapshot waits for canonical config and revision, retaining account catalog', async () => {
  const workspace = { config: { employes_data: [{ id:'same', fdc:'old' }] }, revision: 'old', accounts:[{username:'opaque'}], contract_refresh_required:true }
  let received
  const result = await refreshWorkspaceContracts(workspace, async revision => {
    received = revision
    return { config: { employes_data: [{ id:'same', fdc:'new' }] }, revision:'new' }
  })
  assert.equal(received, 'old')
  assert.equal(result.revision, 'new')
  assert.equal(result.config.employes_data[0].fdc, 'new')
  assert.equal(result.accounts, workspace.accounts)
  assert.equal(workspace.config.employes_data[0].fdc, 'old')
})

test('revision conflict is surfaced without retry or modifying the original workspace', async () => {
  const workspace = { config: { employes_data: [] }, revision: 'old', contract_refresh_required:true }
  let calls = 0
  await assert.rejects(refreshWorkspaceContracts(workspace, async () => { calls++; throw Error('stale revision') }), /stale revision/)
  assert.equal(calls, 1)
  assert.equal(workspace.revision, 'old')
})

 test('background synchronization updates clean views but preserves unsaved or busy work', () => {
  const clean = {dirty:false,busy:false,pending:false}
  assert.equal(canAdoptWorkspaceRevision('old','new',clean),true)
  assert.equal(canAdoptWorkspaceRevision('old','old',clean),false)
  for (const key of ['dirty','busy','pending']) assert.equal(canAdoptWorkspaceRevision('old','new',{...clean,[key]:true}),false)
})
