import { request } from '../api'
import type { Cell, Config, Day, SavedConfig, Workspace } from './types'
import type { ContractInput } from '../admin/api'
import { refreshWorkspaceContracts } from './contractRefresh'

export const attendanceApi = {
  completeContract: (token: string, username: string, month: number, year: number, contract: ContractInput) =>
    request(`/attendance/contracts/${encodeURIComponent(username)}`, token, { method: 'POST', body: JSON.stringify({ ...contract, month, year }) }),
  workspace: async (token: string, signal?: AbortSignal): Promise<Workspace> => {
    const workspace = await request<Workspace>('/attendance/workspace', token, { signal })
    return refreshWorkspaceContracts(workspace, revision => request<SavedConfig>('/attendance/refresh-contracts', token,
      { signal, method: 'POST', body: JSON.stringify({ revision }) }))
  },
  save: (token: string, config: Config, revision: string, deletedIds: string[] = []) => request<SavedConfig>('/attendance/save', token, { method: 'POST', body: JSON.stringify({ config, revision, deleted_ids: deletedIds }) }),
  open: (token: string, username: string, month: number, year: number, revision: string) => request<SavedConfig & { active_id: string }>(`/attendance/open/${encodeURIComponent(username)}`, token,
    { method: 'POST', body: JSON.stringify({ month, year, revision }) }),
  calendar: (token: string, id: string, month: number, year: number, signal?: AbortSignal) => request<{ days: Day[]; read_only?: boolean }>(`/attendance/calendar/${encodeURIComponent(id)}?month=${month}&year=${year}`, token, { signal }),
  apply: (token: string, id: string, month: number, year: number, cells: Cell[], status: string, reason: string, revision: string) =>
    request<SavedConfig>(`/attendance/calendar/${encodeURIComponent(id)}`, token, { method: 'POST', body: JSON.stringify({ month, year, cells, status, reason, revision }) }),
  export: (token: string, ids: string[], month: number, year: number, format: 'original' | 'pdf') =>
    request<{ blob: Blob; filename: string }>('/attendance/export', token, { method: 'POST', body: JSON.stringify({ ids, month, year, format }) }, { blob: true, timeout: 240000 }),
}

export { downloadFile } from '../download'
