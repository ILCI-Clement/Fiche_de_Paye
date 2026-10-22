import { personDisplayName } from '../personDisplayName'
import { useState } from 'react'
import type { MouseEvent, ReactNode } from 'react'
import { Box, Button, Checkbox, Chip, Dialog, DialogActions, DialogContent, DialogTitle, IconButton, ListItemIcon, Menu, MenuItem, Paper, Tooltip, Typography } from '@mui/material'
import { alpha } from '@mui/material/styles'
import MoreHorizRounded from '@mui/icons-material/MoreHorizRounded'
import DeleteOutlineRounded from '@mui/icons-material/DeleteOutlineRounded'
import VisibilityOutlined from '@mui/icons-material/VisibilityOutlined'
import DownloadOutlined from '@mui/icons-material/DownloadOutlined'
import EditOutlined from '@mui/icons-material/EditOutlined'
import ArchiveOutlined from '@mui/icons-material/ArchiveOutlined'
import UnarchiveOutlined from '@mui/icons-material/UnarchiveOutlined'
import { PermanentDeleteDialog } from '../components/PermanentDeleteDialog'
import { PeriodSelector } from '../attendance/PeriodSelector'
import type { Receipt } from './api'
import { periodLabel, receiptPeriod } from './api'

function DocumentAction({ label, disabled, onClick, children }: { label: string; disabled: boolean; onClick: (event: MouseEvent<HTMLButtonElement>) => void; children: ReactNode }) {
  return <Tooltip title={label} arrow describeChild><span style={{ display: 'inline-flex' }}>
    <IconButton aria-label={label} disabled={disabled} onClick={onClick} sx={theme => ({ width: 44, height: 44, borderRadius: 2, color: 'text.secondary', bgcolor: alpha(theme.palette.text.primary, .035), '&:hover': { color: 'primary.main', bgcolor: alpha(theme.palette.primary.main, .12) } })}>{children}</IconButton>
  </span></Tooltip>
}

export type DocumentDensity = 'compact' | 'standard' | 'comfortable'

export function DocumentCard({ receipt: r, admin, canArchive, canDelete, selected, busy, density = 'standard', onSelect, onPreview, onDownload, onArchive, onDelete, onPeriod }: {
  receipt: Receipt; admin: boolean; canArchive: boolean; canDelete: boolean; selected: boolean; busy: boolean;
  density?: DocumentDensity;
  onSelect: () => void; onPreview: () => void; onDownload: () => void; onArchive: () => Promise<boolean>;
  onDelete: (confirmation: string) => Promise<unknown>; onPeriod: (month: number, year: number) => Promise<boolean>
}) {
  const [dialog, setDialog] = useState<'delete' | 'period' | 'archive' | null>(null)
  const [failed, setFailed] = useState(false)
  const [menu, setMenu] = useState<HTMLElement | null>(null)
  const initial = receiptPeriod(r)
  const [period, setPeriod] = useState(initial)
  const compact = density === 'compact'
  const comfortable = density === 'comfortable'
  return <Paper variant="outlined" sx={theme => ({ p: compact ? 2 : comfortable ? 3 : 2.5, minWidth: 0, height: '100%', display: 'flex', flexDirection: 'column', borderRadius: 3, borderColor: selected ? 'primary.main' : alpha(theme.palette.text.primary, .1), transition: 'border-color .15s, box-shadow .15s', '&:hover': { borderColor: alpha(theme.palette.primary.main, .45), boxShadow: `0 6px 20px ${alpha('#000', .08)}` }, '& .MuiChip-root': { border: 0, fontWeight: 400 } })}>
    <Box sx={{ display: 'flex', alignItems: 'start', gap: 1 }}>
      <Checkbox checked={selected} onChange={onSelect} disabled={busy} slotProps={{ input: { 'aria-label': `Sélectionner ${r.original_filename}` } }} sx={{ p: .5 }} />
      <Typography sx={{ fontWeight: 500, fontSize: compact ? 14 : comfortable ? 18 : 16, overflowWrap: 'anywhere', flex: 1 }}>{r.original_filename}<Typography component="span" color="text.secondary" sx={{ display: 'block', fontSize: 12, mt: .5 }}>{periodLabel(initial.month, initial.year)}</Typography></Typography>
    </Box>
    <Box sx={{ my: 1.5, display: 'flex', flexWrap: 'wrap', gap: 1, '& .MuiChip-root': { maxWidth: '100%', ...(compact ? { fontSize: 11, height: 22 } : {}) } }}>
      {r.document_number && <Chip size="small" label={`N° ${r.document_number}`} sx={theme => ({ fontFamily: 'monospace', userSelect: 'text', bgcolor: alpha(theme.palette.primary.main, .12), color: 'primary.main' })} />}
      <Chip size="small" label={r.document_type === 'autre' ? `Autre · ${r.other_type || 'Non précisé'}` : r.document_type === 'transport' ? 'Transport' : 'Non renseigné'} sx={theme => { const color = r.document_type === 'transport' ? theme.palette.info.main : theme.palette.secondary.main; return { bgcolor: alpha(color, .15), color } }} />
      {r.archived && <Chip size="small" label="Archivé" sx={theme => ({ bgcolor: alpha(theme.palette.text.primary, .08), color: 'text.secondary' })} />}
    </Box>
    <Typography variant="body2" color="text.secondary" sx={{ overflowWrap: 'anywhere' }}>Employé : {personDisplayName({ display_name: r.employee_display_name })}</Typography>
    <Typography variant="caption" sx={theme => ({ color: alpha(theme.palette.text.primary, theme.palette.mode === 'dark' ? .45 : .6), display: 'block', overflowWrap: 'anywhere', mt: 1 })}>Déposé par {personDisplayName({ display_name: r.uploaded_by_display_name })} · {r.created_at.slice(0, 10)} · {(r.size_bytes / 1024).toFixed(0)} Ko</Typography>
    {r.archived && r.archived_by_display_name && <Typography variant="caption" color="text.secondary" sx={{ display: 'block', overflowWrap: 'anywhere', mt: .5 }}>Archivé par {personDisplayName({ display_name: r.archived_by_display_name })}</Typography>}
    <Box sx={{ display: 'flex', flexWrap: 'wrap', gap: 1, mt: 'auto', pt: 2 }}>
      <DocumentAction label="Aperçu du document" disabled={busy} onClick={onPreview}><VisibilityOutlined fontSize="small" /></DocumentAction>
      <DocumentAction label="Télécharger le document" disabled={busy} onClick={onDownload}><DownloadOutlined fontSize="small" /></DocumentAction>
      {(admin || canArchive || canDelete) && <Box sx={{ ml: 'auto' }}><DocumentAction label="Autres actions" disabled={busy} onClick={event => { setFailed(false); setMenu(event.currentTarget) }}><MoreHorizRounded fontSize="small" /></DocumentAction></Box>}
    </Box>
    <Menu anchorEl={menu} open={!!menu} onClose={() => setMenu(null)}>
      {admin && <MenuItem disabled={busy} onClick={() => { setMenu(null); setPeriod(initial); setDialog('period') }}><ListItemIcon><EditOutlined fontSize="small" /></ListItemIcon>Modifier le mois</MenuItem>}
      {canArchive && <MenuItem disabled={busy} onClick={() => { setMenu(null); setDialog('archive') }}><ListItemIcon>{r.archived ? <UnarchiveOutlined fontSize="small" /> : <ArchiveOutlined fontSize="small" />}</ListItemIcon>{r.archived ? 'Restaurer' : 'Archiver'}</MenuItem>}
      {canDelete && <MenuItem disabled={busy} sx={{ color: 'error.main' }} onClick={() => { setMenu(null); setDialog('delete') }}><ListItemIcon sx={{ color: 'inherit' }}><DeleteOutlineRounded fontSize="small" /></ListItemIcon>Supprimer</MenuItem>}
    </Menu>
    <Dialog open={dialog === 'period' || dialog === 'archive'} onClose={() => { if (!busy) setDialog(null) }}>
      <DialogTitle>{dialog === 'archive' ? r.archived ? 'Restaurer ce document ?' : 'Archiver ce document ?' : 'Modifier le mois de rattachement'}</DialogTitle>
      <DialogContent>{dialog === 'archive' ? <Typography sx={{ overflowWrap: 'anywhere' }}>{r.original_filename} sera {r.archived ? 'restauré dans le dossier normal' : 'archivé avec ses informations conservées'}.</Typography> : <Box sx={{ pt: 2 }}><PeriodSelector {...period} disabled={busy} onChange={(month, year) => setPeriod({ month, year })} /></Box>}{failed && <Typography color="error" sx={{ mt: 2 }}>Opération refusée. Fermez cette fenêtre pour consulter le détail et actualisez les données.</Typography>}</DialogContent>
      <DialogActions><Button disabled={busy} onClick={() => setDialog(null)}>Annuler</Button><Button variant="contained" disabled={busy} onClick={async () => {
        setFailed(false)
        const ok = dialog === 'archive' ? await onArchive() : await onPeriod(period.month, period.year)
        if (ok) setDialog(null)
        else setFailed(true)
      }}>Confirmer</Button></DialogActions>
    </Dialog>
    {dialog === 'delete' && <PermanentDeleteDialog title="Supprimer définitivement ce document ?" expected={String(r.id)} description={`${r.original_filename} sera retiré du registre et son fichier supprimé.`} onConfirm={onDelete} onDeleted={() => setDialog(null)} onClose={() => setDialog(null)} />}
  </Paper>
}
