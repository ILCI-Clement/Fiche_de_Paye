import { createContext, useContext } from 'react'
import type { WorkspacePage } from './workspacePages'

export interface WorkDestination {
  page: WorkspacePage
  username?: string
  employeeId?: number
  leaveId?: number
  month?: number
  year?: number
  view?: 'editing' | 'signed'
}
export type WorkRequest = WorkDestination & { sequence: number }
export const WorkspaceNavigation = createContext<{ request: WorkRequest | null; navigate: (destination: WorkDestination) => void } | null>(null)
export function useWorkNavigation() { return useContext(WorkspaceNavigation) }
