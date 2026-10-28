import { test } from 'node:test'
import assert from 'node:assert/strict'
import { documentCategoryLabel } from '../src/documents/documentCategory.ts'
import { matchesDocumentSearch } from '../src/documents/documentSearch.ts'

test('custom types use public labels rather than their storage codes', () => {
  const receipt = { document_type: 'c_private-code', category_label: 'Hébergement', other_type: null }
  assert.equal(documentCategoryLabel(receipt), 'Hébergement')
  assert.equal(matchesDocumentSearch(receipt, 'hebergement'), true)
  assert.equal(matchesDocumentSearch(receipt, 'c_private-code'), false)
  assert.equal(matchesDocumentSearch(receipt, 'non renseigné'), false)
})

test('legacy Transport, Autre and unclassified documents keep readable labels', () => {
  assert.equal(documentCategoryLabel({ document_type: 'transport' }), 'Transport')
  assert.equal(documentCategoryLabel({ document_type: 'autre', other_type: 'Taxi' }), 'Autre · Taxi')
  assert.equal(documentCategoryLabel({ document_type: null }), 'Non renseigné')
})
