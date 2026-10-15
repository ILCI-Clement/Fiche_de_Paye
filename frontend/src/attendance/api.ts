import { request } from '../api'
import type { Cell, Config, Day, SavedConfig, Workspace } from './types'

export const attendanceApi = {
  workspace: (token: string, signal?: AbortSignal) => request<Workspace>('/attendance/workspace', token, { signal }),
  save: (token: string, config: Config, revision: string) => request<SavedConfig>('/attendance/save', token, { method: 'POST', body: JSON.stringify({ config, revision }) }),
  open: (token: string, username: string, month: number, year: number, revision: string) => request<SavedConfig & { active_id: string }>(`/attendance/open/${encodeURIComponent(username)}`, token,
    { method: 'POST', body: JSON.stringify({ month, year, revision }) }),
  calendar: (token: string, id: string, month: number, year: number, signal?: AbortSignal) => request<{ days: Day[] }>(`/attendance/calendar/${encodeURIComponent(id)}?month=${month}&year=${year}`, token, { signal }),
  apply: (token: string, id: string, month: number, year: number, cells: Cell[], status: string, reason: string, revision: string) =>
    request<SavedConfig>(`/attendance/calendar/${encodeURIComponent(id)}`, token, { method: 'POST', body: JSON.stringify({ month, year, cells, status, reason, revision }) }),
  export: (token: string, ids: string[], month: number, year: number, format: 'original' | 'pdf') =>
    request<{ blob: Blob; filename: string }>('/attendance/export', token, { method: 'POST', body: JSON.stringify({ ids, month, year, format }) }, { blob: true, timeout: 240000 }),
}

export { downloadFile } from '../download'
