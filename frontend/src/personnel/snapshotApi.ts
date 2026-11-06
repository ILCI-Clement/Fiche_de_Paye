import { request } from '../api'

export interface RegisterSnapshot {
  id: string; source: 'manual' | 'scheduled' | 'catchup'; status: 'running' | 'ready' | 'failed';
  scheduled_for: string | null; captured_at: string | null; created_at: string;
  employee_count: number; row_count: number; error_code: string;
}
export interface SnapshotList {
  settings: { interval_months: number; version: number; next_due: string };
  snapshots: RegisterSnapshot[]; total: number; can_configure: boolean; can_export_excel: boolean; encryption_ready: boolean;
}
const root = '/personnel-snapshots'
const post = (data: unknown) => ({ method: 'POST', body: JSON.stringify(data), cache: 'no-store' as const })
export const snapshotApi = {
  list: (token: string, page: number, signal?: AbortSignal) => request<SnapshotList>(`${root}?page=${page}`, token, { signal, cache: 'no-store' }),
  manual: (token: string, request_id: string) => request<RegisterSnapshot>(root, token, post({ request_id }), { timeout: 120000 }),
  retry: (token: string, id: string) => request<RegisterSnapshot>(`${root}/${id}/retry`, token, post({}), { timeout: 120000 }),
  configure: (token: string, interval_months: number, expected_version: number) => request(root + '/settings', token, { ...post({ interval_months, expected_version }), method: 'PUT' }),
  download: (token: string, id: string, include_nir: boolean, format: 'pdf' | 'xlsx' = 'pdf') => request<{ blob: Blob; filename: string }>(`${root}/${id}/download`, token, post({ include_nir, confirm_sensitive: include_nir, format }), { blob: true, timeout: 120000 }),
}
