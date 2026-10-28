import { request } from '../api'
import type { WorkDestination } from '../workspaceNavigation'

export interface ReadinessIssue { code: string; label: string; domain: 'attendance' | 'contacts' | 'register'; destination: WorkDestination | null }
export interface PersonReadiness { username: string; display_name: string; attendance_ready: boolean; contacts_ready: boolean; issues: ReadinessIssue[] }
export type TaskKind = 'readiness' | 'leave' | 'signature' | 'archive'
export interface WorkTask { id: string; kind: TaskKind; name: string; label: string; destination: WorkDestination | null }
export interface WorkCenter { month: number; year: number; people: PersonReadiness[]; tasks: WorkTask[]; counts: Record<TaskKind, number> }
export const workCenterApi = {
  get: (token: string, month: number, year: number, signal?: AbortSignal) => request<WorkCenter>(`/dashboard/work-center?month=${month}&year=${year}`, token, { signal }),
}
