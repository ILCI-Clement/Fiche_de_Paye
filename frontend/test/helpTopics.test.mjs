import { test } from 'node:test'
import assert from 'node:assert/strict'
import { helpTopicsFor } from '../src/help/helpTopics.ts'

test('guide presents only modules available to the business and system roles', () => {
  const ids = roles => helpTopicsFor({role_tags:roles}).map(topic => topic.id)
  assert.deepEqual(ids(['Employe']),['start','leave','documents','medical','team','profile'])
  assert.ok(ids(['Responsable']).includes('attendance'))
  assert.ok(ids(['Responsable']).includes('dashboard'))
  assert.ok(!ids(['Responsable']).includes('admin'))
  for (const role of ['Admin','SuperAdmin']) {
    assert.ok(ids([role]).includes('personnel'))
    assert.ok(ids([role]).includes('admin'))
  }
})
