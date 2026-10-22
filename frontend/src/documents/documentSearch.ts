import type { Receipt } from './api'

function normalize(value: string) {
  return value.normalize('NFD').replace(/\p{M}/gu, '').toLocaleLowerCase('fr-FR').trim()
}

export function matchesDocumentSearch(receipt: Receipt, query: string) {
  const terms = normalize(query).split(/\s+/).filter(Boolean)
  const text = normalize([
    receipt.original_filename, receipt.document_number, receipt.employee_display_name,
    receipt.uploaded_by_display_name, receipt.archived_by_display_name, receipt.document_type === 'transport' ? 'Transport'
      : receipt.document_type === 'autre' ? 'Autre' : 'Non renseigné', receipt.other_type,
  ].filter(Boolean).join(' '))
  return terms.every(term => text.includes(term))
}
