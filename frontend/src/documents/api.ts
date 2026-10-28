import { request } from '../api'

export interface Receipt {
  id: number; document_number?: string | null; employee_username: string; uploaded_by: string; original_filename: string;
  size_bytes: number; document_type: string | null; other_type: string | null;
  receipt_month: number | null; receipt_year: number | null; archived: boolean; created_at: string;
  employee_display_name?: string; uploaded_by_display_name?: string; archived_by_display_name?: string;
  category_label?: string;
}
export interface DocumentCategory { code: string; label: string; is_active: boolean; is_system: boolean; version: number }
export interface DocumentAccount { username: string; email?: string; full_name?: string | null; display_name?: string }
export interface Upload {
  employee_username: string; original_filename: string; file_b64: string;
  receipt_month: number; receipt_year: number; document_type: string; other_type: string;
}
const root = '/transport-receipts'
export const documentsApi = {
  list: (token: string, signal?: AbortSignal) => request<{ receipts: Receipt[] }>(root, token, { signal }),
  assignees: (token: string, signal?: AbortSignal) => request<{ users: string[]; accounts?: DocumentAccount[] }>(`${root}/assignees`, token, { signal }),
  upload: (token: string, data: Upload) => request(root, token, { method: 'POST', body: JSON.stringify(data) }, { timeout: 120000 }),
  file: (token: string, id: number, signal?: AbortSignal) => request<{ blob: Blob; filename: string }>(`${root}/${id}/file`, token, { signal }, { blob: true, timeout: 60000 }),
  archive: (token: string, id: number, is_archived = true) => request(`${root}/${id}/archive`, token, { method: 'PATCH', body: JSON.stringify({ is_archived }) }),
  remove: (token: string, id: number, confirmation: string) => request<{ cleanup_pending?: boolean; message?: string }>(`${root}/${id}`, token, { method: 'DELETE', body: JSON.stringify({ confirmation }) }),
  period: (token: string, id: number, month: number, year: number) => request(`${root}/${id}/period`, token, { method: 'PATCH', body: JSON.stringify({ receipt_month: month, receipt_year: year }) }),
  category: (token: string, receipt: Receipt, document_type: string, other_type: string) => request(`${root}/${receipt.id}/category`, token, { method: 'PATCH', body: JSON.stringify({ document_type, other_type, expected_document_type: receipt.document_type, expected_other_type: receipt.other_type }) }),
  export: (token: string, ids: number[]) => request<{ blob: Blob; filename: string }>(`${root}/export`, token, { method: 'POST', body: JSON.stringify({ ids }) }, { blob: true, timeout: 240000 }),
}

export const categoryApi = {
  list: (token: string, signal?: AbortSignal) => request<{ categories: DocumentCategory[]; can_manage: boolean }>('/document-categories', token, { signal }),
  create: (token: string, label: string) => request<DocumentCategory>('/document-categories', token, { method: 'POST', body: JSON.stringify({ label }) }),
  update: (token: string, category: DocumentCategory, changes: { label?: string; is_active?: boolean }) => request<DocumentCategory>(`/document-categories/${encodeURIComponent(category.code)}`, token, { method: 'PATCH', body: JSON.stringify({ ...changes, expected_version: category.version }) }),
}

export function receiptPeriod(receipt: Receipt) {
  // Legacy documents without an explicit binding retain their original upload month.
  return { month: receipt.receipt_month ?? Number(receipt.created_at.slice(5, 7)), year: receipt.receipt_year ?? Number(receipt.created_at.slice(0, 4)) }
}
export function periodLabel(month: number, year: number) {
  return new Intl.DateTimeFormat('fr-FR', { month: 'long', year: 'numeric' }).format(new Date(year, month - 1, 1))
}
