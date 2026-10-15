import { useState } from 'react'
import { Box, Button, Checkbox, Chip, Dialog, DialogActions, DialogContent, DialogTitle, Paper, Typography } from '@mui/material'
import { PeriodSelector } from '../attendance/PeriodSelector'
import type { Receipt } from './api'
import { periodLabel, receiptPeriod } from './api'
import { DeleteAction } from '../components/DeleteAction'

export function DocumentCard({ receipt: r, admin, canDelete, selected, busy, onSelect, onPreview, onDownload, onArchive, onDelete, onPeriod }: {
  receipt: Receipt; admin: boolean; canDelete: boolean; selected: boolean; busy: boolean;
  onSelect: () => void; onPreview: () => void; onDownload: () => void; onArchive: () => void;
  onDelete: () => Promise<boolean>; onPeriod: (month: number, year: number) => Promise<boolean>
}) {
  const [dialog, setDialog] = useState<'delete' | 'period' | null>(null)
  const initial = receiptPeriod(r)
  const [period, setPeriod] = useState(initial)
  return <Paper variant="outlined" sx={{ p: 2, minWidth: 0, height: '100%', display: 'flex', flexDirection: 'column' }}>
    <Box sx={{ display: 'flex', alignItems: 'start', gap: 1 }}>
      <Checkbox checked={selected} onChange={onSelect} disabled={busy} slotProps={{ input: { 'aria-label': `Sélectionner ${r.original_filename}` } }} sx={{ p: .5 }} />
      <Typography sx={{ fontWeight: 650, overflowWrap: 'anywhere', flex: 1 }}>{r.original_filename} <Typography component="span" color="primary.main" sx={{ fontSize: 13 }}>· {periodLabel(initial.month, initial.year)}</Typography></Typography>
    </Box>
    <Box sx={{ my: 1, display: 'flex', flexWrap: 'wrap', gap: 1 }}>
      {r.document_number && <Chip size="small" label={`N° ${r.document_number}`} variant="outlined" sx={{ fontFamily: 'monospace', userSelect: 'text', fontWeight: 650 }} />}
      <Chip size="small" label={r.document_type === 'autre' ? `Autre · ${r.other_type}` : r.document_type === 'transport' ? 'Transport' : 'Non renseigné'} color={r.document_type === 'transport' ? 'info' : 'secondary'} variant="outlined" />
      {r.archived && <Chip size="small" label="Archivé" variant="outlined" />}
    </Box>
    <Typography variant="body2" color="text.secondary" sx={{ overflowWrap: 'anywhere' }}>Employé : {r.employee_username.split('@')[0]}</Typography>
    <Typography variant="caption" color="text.secondary" sx={{ display: 'block', overflowWrap: 'anywhere' }}>Déposé par {r.uploaded_by.split('@')[0]} · {r.created_at.slice(0, 10)} · {(r.size_bytes / 1024).toFixed(0)} Ko</Typography>
    <Box sx={{ display: 'flex', flexWrap: 'wrap', gap: .5, mt: 'auto', pt: 1 }}>
      <Button size="small" disabled={busy} onClick={onPreview}>Aperçu</Button>
      <Button size="small" disabled={busy} onClick={onDownload}>Télécharger</Button>
      {admin && <Button size="small" disabled={busy} onClick={() => { setPeriod(initial); setDialog('period') }}>Mois</Button>}
      {admin && !r.archived && <Button size="small" disabled={busy} onClick={onArchive}>Archiver</Button>}
      {canDelete && <DeleteAction label="Supprimer ce document" disabled={busy} onClick={() => setDialog('delete')} />}
    </Box>
    <Dialog open={dialog !== null} onClose={() => { if (!busy) setDialog(null) }}>
      <DialogTitle>{dialog === 'delete' ? 'Supprimer ce document ?' : 'Modifier le mois de rattachement'}</DialogTitle>
      <DialogContent>{dialog === 'delete' ? <Typography sx={{ overflowWrap: 'anywhere' }}>{r.original_filename} sera supprimé définitivement.</Typography> : <Box sx={{ pt: 2 }}><PeriodSelector {...period} disabled={busy} onChange={(month, year) => setPeriod({ month, year })} /></Box>}</DialogContent>
      <DialogActions><Button disabled={busy} onClick={() => setDialog(null)}>Annuler</Button><Button variant="contained" disabled={busy} onClick={async () => {
        const ok = dialog === 'delete' ? await onDelete() : await onPeriod(period.month, period.year)
        if (ok) setDialog(null)
      }}>Confirmer</Button></DialogActions>
    </Dialog>
  </Paper>
}
