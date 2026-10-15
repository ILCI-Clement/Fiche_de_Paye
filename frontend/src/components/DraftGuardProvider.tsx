import { useCallback, useEffect, useMemo, useState } from 'react'
import type { ReactNode } from 'react'
import { Button, Dialog, DialogActions, DialogContent, DialogTitle } from '@mui/material'
import { DraftContext } from '../draftGuard'
import type { Draft } from '../draftRules'
import { draftDecision } from '../draftRules'

export function DraftGuardProvider({ children }: { children: ReactNode }) {
  const [drafts, setDrafts] = useState<Record<string, Draft>>({})
  const [pending, setPending] = useState<{ action: () => void; labels: string[] } | null>(null)
  const set = useCallback((id: string, state?: Draft) => setDrafts(previous => {
    const next = { ...previous }
    if (state) next[id] = state
    else delete next[id]
    return next
  }), [])
  const busy = Object.values(drafts).some(draft => draft.busy)
  const dirty = Object.values(drafts).some(draft => draft.dirty)
  const dirtyLabels = useMemo(() => [...new Set(Object.values(drafts).filter(draft => draft.dirty).map(draft => draft.label))], [drafts])
  useEffect(() => {
    if (!dirty && !busy) return
    const prevent = (event: BeforeUnloadEvent) => { event.preventDefault(); event.returnValue = '' }
    window.addEventListener('beforeunload', prevent)
    return () => window.removeEventListener('beforeunload', prevent)
  }, [dirty, busy])
  const request = useCallback((action: () => void, only?: string, exceptLabel?: string) => {
    const { blocked, labels } = draftDecision(drafts, only, exceptLabel)
    if (blocked) return
    if (labels.length) setPending({ action, labels })
    else action()
  }, [drafts])
  const value = useMemo(() => ({ set, request, busy, dirtyLabels }), [set, request, busy, dirtyLabels])
  return <DraftContext.Provider value={value}>{children}
    <Dialog open={!!pending} onClose={() => setPending(null)}>
      <DialogTitle>Modifications non enregistrées</DialogTitle>
      <DialogContent>{pending?.labels.join(' · ')}. Enregistrez vos modifications avant de continuer. Changer de module conserve les formulaires ouverts ; fermer un formulaire, changer de fiche ou quitter la session peut perdre le brouillon.</DialogContent>
      <DialogActions><Button onClick={() => setPending(null)}>Rester</Button><Button disabled={busy} onClick={() => { const action = pending?.action; setPending(null); action?.() }}>Continuer</Button></DialogActions>
    </Dialog>
  </DraftContext.Provider>
}

