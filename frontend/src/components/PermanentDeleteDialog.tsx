import { useRef, useState } from 'react'
import { Alert, Button, Dialog, DialogActions, DialogContent, DialogTitle, TextField, Typography } from '@mui/material'
import { useDraftGuard } from '../draftGuard'

/** Destructive requests remain subject to live server authorization and reference checks. */
export function PermanentDeleteDialog({ title, description, expected, confirmationLabel, requireReason = false, onConfirm, onDeleted, onClose }: {
  title: string; description: string; expected: string; requireReason?: boolean;
  confirmationLabel?: string;
  onConfirm: (confirmation: string, reason: string) => Promise<unknown>; onDeleted: () => void; onClose: () => void;
}) {
  const [confirmation, setConfirmation] = useState('')
  const [reason, setReason] = useState('')
  const [busy, setBusy] = useState(false)
  const [error, setError] = useState('')
  const lock = useRef(false)
  const close = useDraftGuard(!!confirmation || !!reason, busy, title)
  async function remove() {
    if (lock.current || confirmation !== expected || (requireReason && !reason.trim())) return
    lock.current = true; setBusy(true); setError('')
    try { await onConfirm(confirmation, reason.trim()); onDeleted() }
    catch (failure) { setError(failure instanceof Error ? failure.message : 'Suppression impossible. Actualisez puis réessayez.') }
    finally { lock.current = false; setBusy(false) }
  }
  return <Dialog open fullWidth maxWidth="sm" aria-labelledby="permanent-delete-title" onClose={() => { if (!busy) close(onClose) }}>
    <DialogTitle id="permanent-delete-title">{title}</DialogTitle>
    <DialogContent>
      <Alert severity="warning" sx={{ mb: 2 }}>{description} Cette opération est définitive.</Alert>
      <Typography variant="body2" color="text.secondary" sx={{ mb: 2 }}>Seul un super-administrateur peut effectuer cette suppression. Les références et protections sont vérifiées par le serveur.</Typography>
      <TextField fullWidth label={confirmationLabel || `Saisissez ${expected} pour confirmer`} helperText={confirmationLabel ? `Valeur attendue : ${expected}` : undefined} value={confirmation} disabled={busy} autoComplete="off" onChange={event => setConfirmation(event.target.value)} />
      {requireReason && <TextField fullWidth multiline minRows={2} label="Motif de la suppression" required value={reason} disabled={busy} onChange={event => setReason(event.target.value)} slotProps={{ htmlInput: { maxLength: 500 } }} sx={{ mt: 2 }} />}
      {error && <Alert severity="error" sx={{ mt: 2 }}>{error}</Alert>}
    </DialogContent>
    <DialogActions><Button disabled={busy} onClick={() => close(onClose)}>Annuler</Button><Button color="error" variant="contained" disabled={busy || confirmation !== expected || (requireReason && !reason.trim())} onClick={() => void remove()}>{busy ? 'Suppression…' : 'Supprimer définitivement'}</Button></DialogActions>
  </Dialog>
}
