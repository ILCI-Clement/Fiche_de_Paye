import type { Receipt } from './api'

export function documentCategoryLabel(receipt: Pick<Receipt, 'document_type' | 'other_type' | 'category_label'>) {
  return receipt.category_label || (receipt.document_type === 'transport' ? 'Transport'
    : receipt.document_type === 'autre' ? `Autre · ${receipt.other_type || 'Non précisé'}` : 'Non renseigné')
}
