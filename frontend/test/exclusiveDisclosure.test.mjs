import { test } from 'node:test'
import assert from 'node:assert/strict'
import { createExclusiveDisclosure } from '../src/personnel/exclusiveDisclosure.ts'

test('claiming another number synchronously hides the previous one before starting its request', () => {
  const scope = createExclusiveDisclosure()
  const visible = new Set()
  const first = scope.claim(() => visible.delete('first'))
  visible.add('first')
  const second = scope.claim(() => visible.delete('second'))
  assert.equal(first.isCurrent(), false)
  assert.equal(visible.size, 0)
  visible.add('second')
  assert.deepEqual([...visible], ['second'])
  assert.equal(second.isCurrent(), true)
})

test('a late response cannot reveal a superseded number, even if abort is ignored', async () => {
  const scope = createExclusiveDisclosure()
  let finish
  const response = new Promise(resolve => { finish = resolve })
  let visible = ''
  const first = scope.claim(() => { visible = '' })
  const pending = response.then(() => { if (first.isCurrent()) visible = 'first' })
  const second = scope.claim(() => { visible = '' })
  visible = 'second'
  finish()
  await pending
  assert.equal(visible, 'second')
  assert.equal(second.isCurrent(), true)
})

test('old unmounts and timers cannot release a new disclosure owner', () => {
  const scope = createExclusiveDisclosure()
  const first = scope.claim(() => first.release())
  const second = scope.claim(() => {})
  first.release()
  assert.equal(second.isCurrent(), true)
  second.release()
  assert.equal(second.isCurrent(), false)
})

test('table and detail views share exclusivity even when displaying the same employee', () => {
  const scope = createExclusiveDisclosure()
  let hidden = 0
  const table = scope.claim(() => { hidden++ })
  const detail = scope.claim(() => { hidden++ })
  assert.equal(hidden, 1)
  assert.equal(table.isCurrent(), false)
  detail.release()
  scope.claim(() => {})
  assert.equal(hidden, 1)
})
