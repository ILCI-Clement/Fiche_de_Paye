import { createContext, useContext, useEffect, useId } from 'react'
import type { Draft } from './draftRules'

export const DraftContext = createContext<{
  set: (id: string, state?: Draft) => void
  request: (action: () => void, only?: string, exceptLabel?: string) => void
  busy: boolean
  dirtyLabels: string[]
} | null>(null)

export function useWorkspaceGuard() { return useContext(DraftContext) }

/** Register only state, never form values or credentials. No draft is persisted. */
export function useDraftGuard(dirty: boolean, busy: boolean, label: string) {
  const guard = useWorkspaceGuard()
  const id = useId()
  const set = guard?.set
  useEffect(() => { set?.(id, { dirty, busy, label }); return () => set?.(id) }, [set, id, dirty, busy, label])
  return (action: () => void) => guard ? guard.request(action, id) : action()
}
