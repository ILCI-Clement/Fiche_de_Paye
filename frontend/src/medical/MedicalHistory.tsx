import { useState } from 'react'
import { Alert, Box, Button, Checkbox, CircularProgress, Dialog, DialogActions, DialogContent, DialogTitle, IconButton, TablePagination, Tooltip, Typography } from '@mui/material'
import VisibilityOutlined from '@mui/icons-material/VisibilityOutlined'
import EditOutlined from '@mui/icons-material/EditOutlined'
import HistoryRounded from '@mui/icons-material/HistoryRounded'
import ArchiveOutlined from '@mui/icons-material/ArchiveOutlined'
import UnarchiveOutlined from '@mui/icons-material/UnarchiveOutlined'
import DeleteOutline from '@mui/icons-material/DeleteOutlineRounded'
import { useQuery, useQueryClient } from '@tanstack/react-query'
import dayjs from 'dayjs'
import { periodLabel } from '../documents/api'
import { PermanentDeleteDialog } from '../components/PermanentDeleteDialog'
import { useDraftGuard } from '../draftGuard'
import { medicalApi } from './api'
import { MedicalExport } from './MedicalExport'
import type { MedicalEmployee, MedicalVisit, MedicalPeriod } from './api'

const dateLabel = (date: string) => dayjs(date).format('DD/MM/YYYY')
const rangeLabel = ({ from, to, count }: { from: number; to: number; count: number }) => `${from}–${to} sur ${count}`

export function MedicalHistory({ token, username, admin, superAdmin, archived, period, person, onClose, onEdit, onPreview }: {
  token: string; username: string; admin: boolean; superAdmin: boolean; archived: boolean; period: MedicalPeriod; person: MedicalEmployee;
  onClose: () => void; onEdit: (visit?: MedicalVisit) => void; onPreview: (v: MedicalVisit) => void;
}) {
  const cache = useQueryClient()
  const [selected, setSelected] = useState<number[]>([])
  const [page, setPage] = useState(0), [busy, setBusy] = useState(false), [error, setError] = useState('')
  const [deleting, setDeleting] = useState<MedicalVisit | null>(null), [versions, setVersions] = useState<MedicalVisit | null>(null)
  const close = useDraftGuard(false, busy, 'Visites médicales')
  const query = useQuery({ queryKey: ['medical', username, 'history', person.employee_id, page, period.year, period.month, archived], queryFn: ({ signal }) => medicalApi.history(token, person.employee_id, page + 1, signal, period, archived), retry: false })
  const changed = () => cache.invalidateQueries({ queryKey: ['medical'] })
  async function archive(visit: MedicalVisit) {
    setBusy(true); setError('')
    try { await medicalApi.archive(token, visit, !visit.is_archived); await changed() }
    catch (reason) { setError(reason instanceof Error ? reason.message : 'Archivage impossible.') }
    finally { setBusy(false) }
  }
  return <>
    <Dialog open fullWidth maxWidth="md" onClose={() => close(onClose)} slotProps={{ paper: { sx: { minHeight: 360, maxHeight: '85dvh' } } }}>
      <DialogTitle>{person.full_name} · Visites médicales</DialogTitle>
      <DialogContent dividers>
        <Box sx={{ display: 'flex', alignItems: 'center', justifyContent: 'flex-end' }}>
          <Button size="small" onClick={() => setSelected(selected.length ? [] : (query.data?.visits.filter(v => v.original_filename).map(v => v.id) || []))}>{selected.length ? 'Tout désélectionner' : 'Sélectionner les PDF de la page'}</Button>
          <MedicalExport token={token} period={period} archived={archived} employeeId={person.employee_id} ids={selected} disabled={!selected.length} />
        </Box>
        {(query.error || error) && <Alert severity="error">{error || query.error?.message}</Alert>}
        {query.isLoading && <CircularProgress size={24} />}
        {query.data?.visits.map((visit, index, visits) => <Box key={visit.id}>
          {(index === 0 || visit.visit_date.slice(0,7) !== visits[index-1].visit_date.slice(0,7)) && <Typography component="h3" variant="subtitle2" sx={{ mt: 2, mb: .5, textTransform: 'capitalize', color: 'text.secondary' }}>{periodLabel(Number(visit.visit_date.slice(5,7)), Number(visit.visit_date.slice(0,4)))}</Typography>}
          <Box sx={{ display: 'flex', alignItems: 'center', gap: 1, py: 1.25, borderBottom: 1, borderColor: 'divider', flexWrap: 'wrap' }}>
            <Checkbox size="small" disabled={!visit.original_filename} checked={selected.includes(visit.id)} slotProps={{ input: { 'aria-label': `Sélectionner la visite du ${dateLabel(visit.visit_date)}` } }} onChange={(_, checked) => setSelected(ids => checked ? [...ids, visit.id] : ids.filter(id => id !== visit.id))} />
            <Typography sx={{ minWidth: 90 }}>{dateLabel(visit.visit_date)}</Typography>
            <Typography noWrap variant="body2" color="text.secondary" sx={{ flex: 1, minWidth: 100 }}>{visit.original_filename || 'Attestation à déposer'}</Typography>
            <Box sx={{ display: 'flex', ml: 'auto' }}>
              {visit.original_filename && <Tooltip title="Voir le PDF"><IconButton aria-label="Voir le PDF" onClick={() => onPreview(visit)}><VisibilityOutlined /></IconButton></Tooltip>}
              <Tooltip title="Versions précédentes"><IconButton aria-label="Versions précédentes" onClick={() => setVersions(visit)}><HistoryRounded /></IconButton></Tooltip>
              {!!query.data?.employee.can_edit && !visit.is_archived && <Tooltip title="Corriger la date ou le PDF"><IconButton aria-label="Corriger la visite" disabled={busy} onClick={() => onEdit(visit)}><EditOutlined /></IconButton></Tooltip>}
              {admin && <Tooltip title={visit.is_archived ? 'Restaurer' : 'Archiver'}><IconButton aria-label={visit.is_archived ? 'Restaurer la visite' : 'Archiver la visite'} disabled={busy} onClick={() => void archive(visit)}>{visit.is_archived ? <UnarchiveOutlined /> : <ArchiveOutlined />}</IconButton></Tooltip>}
              {superAdmin && !!visit.is_archived && <Tooltip title="Supprimer définitivement"><IconButton color="error" aria-label="Supprimer définitivement la visite" disabled={busy} onClick={() => setDeleting(visit)}><DeleteOutline /></IconButton></Tooltip>}
            </Box>
          </Box>
        </Box>)}
        {query.data && !query.data.total && <Typography color="text.secondary">Aucune visite pour cette période.</Typography>}
      </DialogContent>
      <TablePagination component="div" count={query.data?.total || 0} rowsPerPage={20} rowsPerPageOptions={[20]} page={page} onPageChange={(_, value) => setPage(value)} labelDisplayedRows={rangeLabel} />
      <DialogActions><Button disabled={busy} onClick={() => close(onClose)}>Fermer</Button>{!!query.data?.employee.can_edit && !archived && <Button variant="contained" onClick={() => onEdit()}>Ajouter une visite</Button>}</DialogActions>
    </Dialog>
    {versions && <MedicalVersions token={token} username={username} visit={versions} onClose={() => setVersions(null)} onPreview={onPreview} />}
    {deleting && <PermanentDeleteDialog title="Supprimer cette visite ?" description="La visite et ses attestations, y compris les anciennes versions, seront supprimées." expected="SUPPRIMER" onClose={() => setDeleting(null)} onConfirm={confirmation => medicalApi.remove(token, deleting, confirmation).then(result => { if (result.cleanup_pending) setError('Visite supprimée ; nettoyage de certains fichiers à vérifier par un administrateur.') })} onDeleted={() => { setDeleting(null); void changed() }} />}
  </>
}

function MedicalVersions({ token, username, visit, onClose, onPreview }: { token: string; username: string; visit: MedicalVisit; onClose: () => void; onPreview: (v: MedicalVisit) => void }) {
  const [page, setPage] = useState(0)
  const query = useQuery({ queryKey: ['medical', username, 'versions', visit.id, page], queryFn: ({ signal }) => medicalApi.versions(token, visit.id, page+1, signal), retry: false })
  return <Dialog open fullWidth maxWidth="sm" onClose={onClose}>
    <DialogTitle>Historique des corrections</DialogTitle>
    <DialogContent dividers>
      {query.error && <Alert severity="error">{query.error.message}</Alert>}
      {query.isLoading && <CircularProgress size={24} />}
      {query.data && <Typography sx={{ mb: 2 }}>Version actuelle · {dateLabel(query.data.current.visit_date)}</Typography>}
      {query.data?.versions.map(previous => <Box key={previous.version} sx={{ display: 'flex', gap: 1, alignItems: 'center', py: 1, borderBottom: 1, borderColor: 'divider' }}>
        <Typography variant="body2" sx={{ flex: 1 }}>Version {previous.version} · {dateLabel(previous.visit_date)}<Typography component="span" sx={{ display: 'block' }} variant="caption" color="text.secondary">Remplacée le {dayjs(previous.created_at).format('DD/MM/YYYY HH:mm')}</Typography></Typography>
        {previous.original_filename && <Tooltip title="Voir l’ancien PDF"><IconButton aria-label={`Voir le PDF version ${previous.version}`} onClick={() => onPreview({ ...visit, ...previous, revision: previous.version })}><VisibilityOutlined /></IconButton></Tooltip>}
      </Box>)}
      {query.data && !query.data.total && <Typography color="text.secondary">Aucune correction enregistrée.</Typography>}
    </DialogContent>
    <TablePagination component="div" count={query.data?.total || 0} rowsPerPage={20} rowsPerPageOptions={[20]} page={page} onPageChange={(_, value) => setPage(value)} labelDisplayedRows={rangeLabel} />
    <DialogActions><Button onClick={onClose}>Fermer</Button></DialogActions>
  </Dialog>
}
