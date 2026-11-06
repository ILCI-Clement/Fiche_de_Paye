import { request } from '../api'

export interface MedicalEmployee {
  employee_id: number; full_name: string; is_archived: boolean; departments: string[];
  latest_id: number | null; visit_date: string | null; original_filename: string | null; visit_count: number;
}
export interface MedicalVisit { id: number; visit_date: string; original_filename: string | null; size_bytes: number; version: number; created_at: string; is_archived: boolean; revision?: number }
export interface PdfInput { original_filename: string; file_b64: string }
export interface MedicalPeriod { year: number; month: number | null }
export function medicalPeriodQuery(period?: MedicalPeriod) {
  return period ? `&year=${period.year}${period.month ? `&month=${period.month}` : ''}` : ''
}
export const medicalApi = {
  employees: (token: string, q: string, page: number, archived: boolean, signal?: AbortSignal, period?: MedicalPeriod, mine = false) => request<{ employees: MedicalEmployee[]; total: number; years: number[]; export_count: number }>(`/medical-visits/employees?${new URLSearchParams({ q, page: String(page), archived: String(archived), mine: String(mine) })}${medicalPeriodQuery(period)}`, token, { signal }),
  history: (token: string, id: number, page: number, signal?: AbortSignal, period?: MedicalPeriod, archived = false) => request<{ visits: MedicalVisit[]; total: number; employee: { employee_id: number; full_name: string; is_archived: boolean; can_edit: boolean } }>(`/medical-visits/employees/${id}?page=${page}&archived=${archived}${medicalPeriodQuery(period)}`, token, { signal }),
  export: (token: string, period: MedicalPeriod, q: string, archived: boolean, employee_id?: number, format: 'zip' = 'zip', visit_ids?: number[], all_history = false) => request<{ blob: Blob; filename: string }>('/medical-visits/export', token, { method: 'POST', body: JSON.stringify({ ...period, q, archived, employee_id, format, visit_ids, all_history }) }, { blob: true, timeout: 120000 }),
  record: (token: string, id: number, signal?: AbortSignal) => request<MedicalVisit>(`/medical-visits/${id}`, token, { signal }),
  create: (token: string, employee_id: number, visit_date: string, pdf: PdfInput | null) => request('/medical-visits', token, { method: 'POST', body: JSON.stringify({ employee_id, visit_date, pdf }) }, { timeout: 60000 }),
  attach: (token: string, visit: MedicalVisit, pdf: PdfInput) => request(`/medical-visits/${visit.id}/file`, token, { method: 'POST', body: JSON.stringify({ ...pdf, expected_version: visit.version }) }, { timeout: 60000 }),
  correct: (token: string, visit: MedicalVisit, visit_date: string, pdf: PdfInput | null) => request(`/medical-visits/${visit.id}`, token, { method: 'PUT', body: JSON.stringify({ expected_version: visit.version, visit_date, pdf }) }, { timeout: 60000 }),
  archive: (token: string, visit: MedicalVisit, is_archived: boolean) => request(`/medical-visits/${visit.id}/archive`, token, { method: 'PATCH', body: JSON.stringify({ expected_version: visit.version, is_archived }) }),
  remove: (token: string, visit: MedicalVisit, confirmation: string) => request<{ cleanup_pending: boolean }>(`/medical-visits/${visit.id}`, token, { method: 'DELETE', body: JSON.stringify({ expected_version: visit.version, confirmation }) }),
  versions: (token: string, id: number, page: number, signal?: AbortSignal) => request<{ current: MedicalVisit; versions: Omit<MedicalVisit, 'id' | 'is_archived'>[]; total: number }>(`/medical-visits/${id}/versions?page=${page}`, token, { signal }),
  file: (token: string, id: number, signal?: AbortSignal, revision?: number) => request<{ blob: Blob; filename: string }>(revision == null ? `/medical-visits/${id}/file` : `/medical-visits/${id}/versions/${revision}/file`, token, { signal }, { blob: true, timeout: 60000 }),
}

export async function readPdf(file: File): Promise<PdfInput> {
  if (!file.size || file.size > 10 * 1024 * 1024 || !/\.pdf$/i.test(file.name)) throw new Error('Choisissez un PDF de 10 Mo maximum.')
  return new Promise((resolve, reject) => {
    const reader = new FileReader()
    reader.onload = () => resolve({ original_filename: file.name, file_b64: String(reader.result).split(',')[1] })
    reader.onerror = () => reject(new Error('Impossible de lire le fichier.'))
    reader.readAsDataURL(file)
  })
}
