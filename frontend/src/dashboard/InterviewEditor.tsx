import { useDraftGuard } from '../draftGuard'
import { useState } from 'react'
import { Box, Button, Checkbox, Dialog, DialogActions, DialogContent, DialogTitle, FormControlLabel, TextField } from '@mui/material'
import type { Interview } from './api'
export function InterviewEditor({ item, busy, onClose, save }: { item: Interview; busy: boolean; onClose: () => void; save: (data: { due_date: string | null; notes: string; completed: boolean }) => Promise<boolean> }) {
  const [due, setDue] = useState(item.due_date || '')
  const [notes, setNotes] = useState(item.notes || '')
  const [completed, setCompleted] = useState(Boolean(item.completed))
  const close = useDraftGuard(due !== (item.due_date || '') || notes !== (item.notes || '') || completed !== Boolean(item.completed), busy, 'Entretien annuel')
  return <Dialog open onClose={() => { if (!busy) close(onClose) }} fullWidth maxWidth="sm"><DialogTitle>Entretien · {item.employee_username}</DialogTitle>
    <Box component="form" onSubmit={async e => { e.preventDefault(); if (await save({ due_date: due || null, notes, completed })) onClose() }}>
      <DialogContent sx={{ display: 'grid', gap: 2 }}><TextField label="Échéance" type="date" value={due} disabled={busy} onChange={e => setDue(e.target.value)} slotProps={{ inputLabel: { shrink: true } }} helperText="Une date laissée vide conserve l’échéance existante." /><TextField label="Remarques" multiline minRows={3} value={notes} disabled={busy} onChange={e => setNotes(e.target.value)} slotProps={{ htmlInput: { maxLength: 5000 } }} /><FormControlLabel control={<Checkbox checked={completed} disabled={busy} onChange={e => setCompleted(e.target.checked)} />} label="Entretien terminé" /></DialogContent>
      <DialogActions><Button disabled={busy} onClick={() => close(onClose)}>Annuler</Button><Button type="submit" variant="contained" disabled={busy}>Enregistrer</Button></DialogActions>
    </Box></Dialog>
}
