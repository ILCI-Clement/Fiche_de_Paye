import { FrenchDateField } from './FrenchDateField'
import { useDraftGuard } from '../draftGuard'
import { useState } from 'react'
import { Alert, Box, Button, Dialog, DialogActions, DialogContent, DialogTitle, MenuItem, Stack, TextField, Typography } from '@mui/material'
import { useMutation } from '@tanstack/react-query'
import { api } from '../api'
import type { LeaveInput, Session, Slot } from '../types'
import { slotLabels } from '../leaveFormat'

function today() {
  const now = new Date()
  return `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')}`
}

export function LeaveForm({ token, session, onClose, onSaved }: {
  token: string; session: Session; onClose: () => void; onSaved: () => void
}) {
  const [input, setInput] = useState<LeaveInput>({ start_date: today(), end_date: today(), start_slot: 'full', end_slot: 'full', note: '' })
  const [baseline] = useState(() => JSON.stringify(input))
  const sameDay = input.start_date === input.end_date
  const valid = /^\d{4}-\d{2}-\d{2}$/.test(input.start_date) && /^\d{4}-\d{2}-\d{2}$/.test(input.end_date) && input.end_date >= input.start_date
  const save = useMutation({ mutationFn: () => api.create(token, input), onSuccess: onSaved })
  const close = useDraftGuard(JSON.stringify(input) !== baseline, save.isPending, 'Demande de congé')
  function setDate(field: 'start_date' | 'end_date', value: string) {
    setInput(previous => {
      const next = { ...previous, [field]: value }
      if (next.start_date === next.end_date) next.end_slot = next.start_slot
      return next
    })
  }
  return <Dialog open onClose={save.isPending ? undefined : () => close(onClose)} fullWidth maxWidth="sm">
    <Box component="form" onSubmit={event => { event.preventDefault(); if (valid && !save.isPending) save.mutate() }}>
      <DialogTitle sx={{ pb: 1 }}>Nouvelle demande de congé</DialogTitle>
      <DialogContent>
        <Typography color="text.secondary" sx={{ mb: 3 }}>Congés payés · Pour {session.name}</Typography>
        {save.error && <Alert severity="error" sx={{ mb: 2 }}>{save.error.message}</Alert>}
        <Stack spacing={2.5} sx={{ pt: 1 }}>
          <Box className="form-row">
            <FrenchDateField label="Date de début" required value={input.start_date} disabled={save.isPending}
              onChange={value => setDate('start_date', value)} fullWidth />
            <TextField select label="Période de début" value={input.start_slot} disabled={save.isPending} fullWidth onChange={event => {
              const slot = event.target.value as Slot
              setInput(previous => ({ ...previous, start_slot: slot, ...(sameDay ? { end_slot: slot } : {}) }))
            }}>{Object.entries(slotLabels).map(([value, label]) => <MenuItem key={value} value={value}>{label}</MenuItem>)}</TextField>
          </Box>
          <Box className="form-row">
            <FrenchDateField label="Date de fin" required value={input.end_date} disabled={save.isPending}
              error={Boolean(input.end_date && input.end_date < input.start_date)}
              helperText={input.end_date < input.start_date ? 'La fin ne peut pas précéder le début.' : undefined}
              onChange={value => setDate('end_date', value)} fullWidth />
            <TextField select label="Période de fin" value={input.end_slot} disabled={sameDay || save.isPending} fullWidth
              onChange={event => setInput(previous => ({ ...previous, end_slot: event.target.value as Slot }))}>
              {Object.entries(slotLabels).map(([value, label]) => <MenuItem key={value} value={value}>{label}</MenuItem>)}
            </TextField>
          </Box>
          {!sameDay && <Typography variant="caption" color="text.secondary">Les dates intermédiaires sont incluses en journées entières.</Typography>}
          <TextField label="Note (facultative)" placeholder="Ajoutez une précision pour votre responsable…" multiline minRows={3}
            value={input.note} disabled={save.isPending} onChange={event => setInput(previous => ({ ...previous, note: event.target.value }))}
            slotProps={{ htmlInput: { maxLength: 2000 } }} helperText={`${input.note.length}/2000 caractères`} />
          <Alert severity="info" icon={false}>Votre demande sera traitée par <strong>{session.approver_name}</strong>.</Alert>
        </Stack>
      </DialogContent>
      <DialogActions sx={{ px: 3, pb: 3 }}>
        <Button onClick={() => close(onClose)} disabled={save.isPending} color="inherit">Annuler</Button>
        <Button type="submit" variant="contained" disabled={!valid || save.isPending}>{save.isPending ? 'Envoi…' : 'Envoyer la demande'}</Button>
      </DialogActions>
    </Box>
  </Dialog>
}
