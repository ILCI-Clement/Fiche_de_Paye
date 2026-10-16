import { useState } from 'react'
import type { ReactNode } from 'react'
import { Box, Button, Checkbox, Chip, Dialog, DialogActions, DialogContent, DialogTitle, IconButton, Paper, Tooltip, Typography } from '@mui/material'
import VisibilityOutlined from '@mui/icons-material/VisibilityOutlined'
import DownloadOutlined from '@mui/icons-material/DownloadOutlined'
import EditOutlined from '@mui/icons-material/EditOutlined'
import ArchiveOutlined from '@mui/icons-material/ArchiveOutlined'
import { PeriodSelector } from '../attendance/PeriodSelector'
import type { Receipt } from './api'
import { periodLabel, receiptPeriod } from './api'
import { DeleteAction } from '../components/DeleteAction'

function DocumentAction({ label, disabled, onClick, children }: { label: string; disabled: boolean; onClick: () => void; children: ReactNode }) {
  return <Tooltip title={label} arrow describeChild><span style={{ display: 'inline-flex' }}>
    <IconButton color="primary" aria-label={label} disabled={disabled} onClick={onClick} sx={{ width: 44, height: 44 }}>{children}</IconButton>
  </span></Tooltip>
}

export type DocumentDensity = 'compact' | 'standard' | 'comfortable'

export function DocumentCard({ receipt: r, admin, canDelete, selected, busy, density = 'standard', onSelect, onPreview, onDownload, onArchive, onDelete, onPeriod }: {
  receipt: Receipt; admin: boolean; canDelete: boolean; selected: boolean; busy: boolean;
  density?: DocumentDensity;
  onSelect: () => void; onPreview: () => void; onDownload: () => void; onArchive: () => void;
  onDelete: () => Promise<boolean>; onPeriod: (month: number, year: number) => Promise<boolean>
}) {
  const [dialog, setDialog] = useState<'delete' | 'period' | null>(null)
  const initial = receiptPeriod(r)
  const [period, setPeriod] = useState(initial)
  const compact = density === 'compact'
  const comfortable = density === 'comfortable'
  return <Paper variant="outlined" sx={{ p: compact ? 1.25 : comfortable ? 3 : 2, minWidth: 0, height: '100%', display: 'flex', flexDirection: 'column' }}>
    <Box sx={{ display: 'flex', alignItems: 'start', gap: 1 }}>
      <Checkbox checked={selected} onChange={onSelect} disabled={busy} slotProps={{ input: { 'aria-label': `Sélectionner ${r.original_filename}` } }} sx={{ p: .5 }} />
      <Typography sx={{ fontWeight: 650, fontSize: compact ? 14 : comfortable ? 18 : 16, overflowWrap: 'anywhere', flex: 1 }}>{r.original_filename} <Typography component="span" color="primary.main" sx={{ fontSize: compact ? 12 : 13 }}>· {periodLabel(initial.month, initial.year)}</Typography></Typography>
    </Box>
    <Box sx={{ my: compact ? .5 : 1, display: 'flex', flexWrap: 'wrap', gap: compact ? .5 : 1, '& .MuiChip-root': { maxWidth: '100%', ...(compact ? { fontSize: 11, height: 22 } : {}) } }}>
      {r.document_number && <Chip size="small" label={`N° ${r.document_number}`} variant="outlined" sx={{ fontFamily: 'monospace', userSelect: 'text', fontWeight: 650 }} />}
      <Chip size="small" label={r.document_type === 'autre' ? `Autre · ${r.other_type}` : r.document_type === 'transport' ? 'Transport' : 'Non renseigné'} color={r.document_type === 'transport' ? 'info' : 'secondary'} variant="outlined" />
      {r.archived && <Chip size="small" label="Archivé" variant="outlined" />}
    </Box>
    <Typography variant="body2" color="text.secondary" sx={{ overflowWrap: 'anywhere' }}>Employé : {r.employee_username.split('@')[0]}</Typography>
    <Typography variant="caption" color="text.secondary" sx={{ display: 'block', overflowWrap: 'anywhere' }}>Déposé par {r.uploaded_by.split('@')[0]} · {r.created_at.slice(0, 10)} · {(r.size_bytes / 1024).toFixed(0)} Ko</Typography>
    <Box sx={{ display: 'flex', flexWrap: 'wrap', gap: compact ? 0 : .5, mt: 'auto', pt: compact ? .5 : 1 }}>
      <DocumentAction label="Aperçu du document" disabled={busy} onClick={onPreview}><VisibilityOutlined fontSize="small" /></DocumentAction>
      <DocumentAction label="Télécharger le document" disabled={busy} onClick={onDownload}><DownloadOutlined fontSize="small" /></DocumentAction>
      {admin && <DocumentAction label="Modifier le mois de rattachement" disabled={busy} onClick={() => { setPeriod(initial); setDialog('period') }}><EditOutlined fontSize="small" /></DocumentAction>}
      {admin && !r.archived && <DocumentAction label="Archiver le document" disabled={busy} onClick={onArchive}><ArchiveOutlined fontSize="small" /></DocumentAction>}
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
