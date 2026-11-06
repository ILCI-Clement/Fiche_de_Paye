import { test } from 'node:test'
import assert from 'node:assert/strict'
import { linkedLeave } from '../src/leaveLink.ts'

test('mail link selects exact request without granting approval or carrying session credentials', () => {
  assert.equal(linkedLeave('?leave=42'),42)
  for (const search of ['', '?leave=0','?leave=-2','?leave=abc','?leave=1.2','?leave=99999999999999999999','?token=secret']) {
    assert.equal(linkedLeave(search),null)
  }
})
