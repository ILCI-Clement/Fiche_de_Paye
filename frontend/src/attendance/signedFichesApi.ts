import { request } from '../api'
import type { FileAction, FileBatchResult } from '../fileBatch'

export interface SignedFiche {
  id: string; document_id: string | null; employee_name: string; fiche_month: number; fiche_year: number;
  is_archived: boolean; mode: string; status: string; created_at: string; completed_at: string | null; archived_at: string | null;
  file_bytes: number | null; last_error: string | null;
}
export interface SignedFiches {
  items: SignedFiche[]; total: number; page: number; can_export: boolean; can_delete: boolean; automatic_archive: boolean;
}
export const signedFichesApi = {
  batch: (token: string, ids: string[], action: FileAction, expected_archived: boolean, confirmation?: string) => request<FileBatchResult>('/signed-fiches/batch', token, { method: 'POST', body: JSON.stringify({ ids, action, expected_archived, confirmation }) }),
  list: (token: string, month: number | undefined, year: number, page: number, mode: string, username?: string, signal?: AbortSignal, archived = false) =>
    request<SignedFiches>(`/signed-fiches?archived=${archived}&year=${year}${month ? `&month=${month}` : ''}&page=${page}&mode=${mode}${username ? `&username=${encodeURIComponent(username)}` : ''}`, token, { signal }),
  file: (token: string, id: string, signal?: AbortSignal) => request<{ blob: Blob; filename: string }>(`/signed-fiches/${encodeURIComponent(id)}/file`, token, { signal }, { blob: true, timeout: 60000 }),
  export: (token: string, month: number, year: number, mode: string) => request<{ blob: Blob; filename: string }>(`/signed-fiches/export.zip?month=${month}&year=${year}&mode=${mode}`, token, undefined, { blob: true, timeout: 240000 }),
  exportYear: (token: string, year: number, mode: string) => request<{ blob: Blob; filename: string }>(`/signed-fiches/export.zip?year=${year}&mode=${mode}`, token, undefined, { blob: true, timeout: 240000 }),
  retry: (token: string, id: string) => request(`/signed-fiches/${encodeURIComponent(id)}/retry`, token, { method: 'POST' }),
  delete: (token: string, id: string, confirmation: string) => request<{ status: string; cleanup_pending: boolean }>(`/signed-fiches/${encodeURIComponent(id)}`, token, { method: 'DELETE', body: JSON.stringify({ confirmation }) }),
}

export const signedStatus: Record<string, string> = {
  pending: 'En cours d’envoi', uncertain: 'Envoi à vérifier', sent: 'Signatures en attente',
  partially_signed: 'Une signature reçue', completed: 'Signé · récupération en cours', ready: 'Signé',
  declined: 'Refusé', expired: 'Expiré', cancelled: 'Annulé',
}
