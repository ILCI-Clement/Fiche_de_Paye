import { test } from 'node:test'
import assert from 'node:assert/strict'
import { ApiError, request } from '../src/api.ts'

test('missing contract is recoverable by its explicit response requirement', async context => {
  context.mock.method(globalThis, 'fetch', async () => new Response(JSON.stringify({ detail: 'Contrat requis.' }), {
    status: 409, headers: { 'X-Attendance-Requirement': 'contract' },
  }))
  await assert.rejects(request('/attendance/open/synthetic', 'synthetic-token'), error =>
    error instanceof ApiError && error.status === 409 && error.requirement === 'contract' && error.message === 'Contrat requis.')
})

test('revision and archive conflicts never open the contract completion flow', async context => {
  context.mock.method(globalThis, 'fetch', async () => new Response(JSON.stringify({ detail: 'Configuration modifiée.' }), { status: 409 }))
  await assert.rejects(request('/attendance/save', 'synthetic-token'), error =>
    error instanceof ApiError && error.status === 409 && error.requirement === undefined)
})

test('contract error handling preserves the existing login cooldown metadata', async context => {
  context.mock.method(globalThis, 'fetch', async () => new Response(JSON.stringify({ detail: 'Réessayez plus tard.' }), {
    status: 429, headers: { 'Retry-After': '600' },
  }))
  await assert.rejects(request('/login', null), error =>
    error instanceof ApiError && error.retryAfterSeconds === 600 && error.requirement === undefined)
})
