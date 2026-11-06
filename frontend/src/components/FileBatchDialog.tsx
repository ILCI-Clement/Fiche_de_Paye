import { useRef, useState } from 'react'
import { Alert, Box, Button, Dialog, DialogActions, DialogContent, DialogTitle, Typography } from '@mui/material'
import { PermanentDeleteDialog } from './PermanentDeleteDialog'
import type { FileAction } from '../fileBatch'
import { fileActionLabels } from '../fileBatch'

/** Freeze and show exactly the selected files; submitting never reuses live selection. */
export function FileBatchDialog({ action, items, onConfirm, onClose }: {
  action: FileAction; items: string[]; onConfirm: (confirmation?: string) => Promise<unknown>; onClose: () => void
}) {
  const [busy, setBusy] = useState(false)
  const [error, setError] = useState('')
  const lock = useRef(false)
  const title = `${fileActionLabels[action]} ${items.length} fichier(s) ?`
  const list = <Box component="ul" sx={{ maxHeight: 160, overflowY: 'auto', pl: 2.5, mb: 2, fontSize: 13 }}>
    {items.map((item, i) => <Box component="li" key={i} sx={{ overflowWrap: 'anywhere', mb: .5 }}>{item}</Box>)}
  </Box>
  if (action === 'delete') return <PermanentDeleteDialog title={title} expected="SUPPRIMER"
    description="Les fichiers sélectionnés et leurs lignes seront supprimés de cette application."
    onConfirm={confirmation => onConfirm(confirmation)} onDeleted={onClose} onClose={onClose}>{list}</PermanentDeleteDialog>
  async function submit() {
    if (lock.current) return
    lock.current = true; setBusy(true); setError('')
    try { await onConfirm(); onClose() }
    catch (reason) { setError(reason instanceof Error ? reason.message : 'Opération impossible.') }
    finally { lock.current = false; setBusy(false) }
  }
  return <Dialog open fullWidth maxWidth="sm" onClose={busy ? undefined : onClose}>
    <DialogTitle>{title}</DialogTitle>
    <DialogContent><Typography variant="body2">
      {action === 'archive' ? 'Les fichiers passent dans Archives. Ils restent consultables et peuvent être restaurés.' : 'Les fichiers reviennent dans le dossier normal.'}
    </Typography>{list}{error && <Alert severity="error">{error}</Alert>}</DialogContent>
    <DialogActions><Button disabled={busy} onClick={onClose}>Annuler</Button>
      <Button variant="contained" disabled={busy} onClick={() => void submit()}>{busy ? 'En cours…' : fileActionLabels[action]}</Button>
    </DialogActions>
  </Dialog>
}
