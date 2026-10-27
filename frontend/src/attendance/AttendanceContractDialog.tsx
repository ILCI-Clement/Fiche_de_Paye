import { useState } from 'react'
import { Alert, Button, Dialog, DialogActions, DialogContent, DialogTitle, Typography } from '@mui/material'
import { attendanceApi } from './api'
import { ContractForm } from '../admin/ContractEditor'
import { personDisplayName } from '../personDisplayName'
import { periodLabel } from '../documents/api'
import type { Account } from './types'

/** A missing contract is a recoverable setup step, never an invented date. */
export function AttendanceContractDialog({ token, account, month, year, canEdit, onClose, onSaved }: {
  token: string; account: Account; month: number; year: number; canEdit: boolean;
  onClose: () => void; onSaved: () => Promise<void>;
}) {
  const [busy, setBusy] = useState(false)
  const [error, setError] = useState('')
  const [saved, setSaved] = useState(false)
  async function openSaved() {
    setBusy(true); setError('')
    try { await onSaved() }
    catch (failure) { setError(failure instanceof Error ? failure.message : 'Contrat enregistré. Réessayez l’ouverture de la fiche.') }
    finally { setBusy(false) }
  }
  return <Dialog open fullWidth maxWidth="sm" onClose={() => { if (!busy) onClose() }}>
    <DialogTitle>Compléter le contrat</DialogTitle>
    <DialogContent>
      <Typography sx={{ mb: 1 }}>{personDisplayName(account)} · {periodLabel(month, year)}</Typography>
      <Typography variant="body2" color="text.secondary" sx={{ mb: 2 }}>
        Aucun contrat ne couvre ce mois. Les dates définissent les jours à remplir sur la fiche.
        {!canEdit && ' Un Admin doit compléter ou vérifier le contrat dans Administration.'}
      </Typography>
      {error && <Alert severity="error" sx={{ mb: 2 }}>{error}</Alert>}
      {saved ? <Button disabled={busy} variant="contained" onClick={() => void openSaved()}>Ouvrir la fiche</Button> : canEdit && <ContractForm busy={busy} disabled={busy} save={async data => {
        const first = `${year}-${String(month).padStart(2, '0')}-01`
        const last = `${year}-${String(month).padStart(2, '0')}-${new Date(year, month, 0).getDate()}`
        if (!data.start_date || data.start_date > last || (data.end_date && data.end_date < first)) {
          setError('Le contrat doit couvrir le mois sélectionné.'); return false
        }
        setBusy(true); setError('')
        try {
          await attendanceApi.completeContract(token, account.username, month, year, data)
          setSaved(true)
        } catch (failure) { setError(failure instanceof Error ? failure.message : 'Contrat non enregistré.'); return false }
        finally { setBusy(false) }
        await openSaved()
        return true
      }} />}
    </DialogContent>
    <DialogActions><Button disabled={busy} onClick={onClose}>Fermer</Button></DialogActions>
  </Dialog>
}
