import { test } from 'node:test'
import assert from 'node:assert/strict'
import { draftDecision } from '../src/draftRules.ts'

test('clean forms navigate without confirmation', () => {
  assert.deepEqual(draftDecision({ a: { dirty: false, busy: false, label: 'A' } }), { blocked: false, labels: [] })
})
test('logout collects dirty forms from hidden modules', () => {
  assert.deepEqual(draftDecision({ a: { dirty: true, busy: false, label: 'A' }, b: { dirty: true, busy: false, label: 'B' } }), { blocked: false, labels: ['A', 'B'] })
})
test('an active operation blocks navigation even with a clean form', () => {
  assert.equal(draftDecision({ a: { dirty: false, busy: true, label: 'Upload' } }).blocked, true)
})
test('local close only checks its own form', () => {
  assert.deepEqual(draftDecision({ a: { dirty: true, busy: false, label: 'A' }, b: { dirty: false, busy: false, label: 'B' } }, 'b'), { blocked: false, labels: [] })
})
test('removed forms do not leave stale warnings', () => {
  assert.deepEqual(draftDecision({}, 'removed'), { blocked: false, labels: [] })
})
test('duplicate labels are shown once', () => {
  assert.deepEqual(draftDecision({ a: { dirty: true, busy: false, label: 'A' }, b: { dirty: true, busy: false, label: 'A' } }).labels, ['A'])
})
test('saving profile still warns about other drafts before renewing session', () => {
  const result = draftDecision({ profile: { dirty: true, busy: false, label: 'Profile' }, attendance: { dirty: true, busy: false, label: 'Attendance' } }, undefined, 'Profile')
  assert.deepEqual(result, { blocked: false, labels: ['Attendance'] })
})
