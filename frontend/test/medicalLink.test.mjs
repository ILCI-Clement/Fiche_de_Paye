import { test } from 'node:test'
import assert from 'node:assert/strict'
import { linkedMedical } from '../src/medical/medicalLink.ts'
test('medical Excel links select a record without accepting credentials or invalid IDs', () => {
  assert.equal(linkedMedical('?medical=42'),42)
  for (const query of ['', '?medical=0', '?medical=-1', '?medical=1.5', '?medical=1e3', '?medical=99999999999999999999', '?token=secret']) assert.equal(linkedMedical(query),null)
})
