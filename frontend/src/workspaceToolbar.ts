import { createContext, useContext } from 'react'

export interface WorkspaceToolbarSlot {
  active: boolean
  target: HTMLElement | null
  inlineTarget: HTMLElement | null
}

/** Each persistent page receives its own visibility flag and the shared header host. */
export const WorkspaceToolbarContext = createContext<WorkspaceToolbarSlot | null>(null)
export function useWorkspaceToolbar() { return useContext(WorkspaceToolbarContext) }
