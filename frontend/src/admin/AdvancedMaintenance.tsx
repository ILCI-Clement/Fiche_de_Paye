import { useRef, useState } from 'react'
import { Alert, Box, Button, CircularProgress, Dialog, DialogActions, DialogContent, DialogTitle, IconButton, MenuItem, Snackbar, Table, TableBody, TableCell, TableContainer, TableHead, TablePagination, TableRow, TextField, ToggleButton, ToggleButtonGroup, Tooltip, Typography } from '@mui/material'
import ArchiveOutlined from '@mui/icons-material/ArchiveOutlined'
import UnarchiveOutlined from '@mui/icons-material/UnarchiveOutlined'
import TuneRounded from '@mui/icons-material/TuneRounded'
import { useQuery, useQueryClient } from '@tanstack/react-query'
import { request } from '../api'
import type { Session } from '../types'
import { isAdmin, isSuperAdmin } from '../rolePolicy'
import { useDraftGuard, useWorkspaceGuard } from '../draftGuard'
import { DeleteAction } from '../components/DeleteAction'
import { PermanentDeleteDialog } from '../components/PermanentDeleteDialog'

type Kind = 'group' | 'leave_request' | 'annual_interview'
interface Resource { id: string; label: string; detail: string; is_archived: boolean }
const labels: [Kind, string][] = [['group', 'Départements'], ['leave_request', 'Congés'], ['annual_interview', 'Entretiens']]
const resourcePath = (kind: Kind, id?: string) => `/maintenance/resources/${kind}${id === undefined ? '' : `/${encodeURIComponent(id)}`}`

export function AdvancedMaintenance({ token, session }: { token: string; session: Session }) {
  const [open, setOpen] = useState(false)
  const [kind, setKind] = useState<Kind>('group')
  const [archive, setArchive] = useState<'active' | 'archived'>('active')
  const [page, setPage] = useState(0)
  const [busy, setBusy] = useState(false)
  const lock = useRef(false)
  const [error, setError] = useState('')
  const [message, setMessage] = useState('')
  const [archiving, setArchiving] = useState<Resource | null>(null)
  const [deleting, setDeleting] = useState<Resource | null>(null)
  const client = useQueryClient()
  const guard = useWorkspaceGuard()
  useDraftGuard(false, busy, 'Gestion avancée')
  const list = useQuery({ queryKey: ['maintenance', session.username, kind, archive, page], queryFn: ({ signal }) => request<{ items: Resource[]; total: number; page: number; page_size: number }>(`${resourcePath(kind)}?archive=${archive}&page=${page + 1}&page_size=25`, token, { signal }), enabled: open && isAdmin(session), retry: false })
  const lastPage = list.data ? Math.max(0, Math.ceil(list.data.total / 25) - 1) : page
  if (page > lastPage) setPage(lastPage)
  async function updateArchive() {
    if (!archiving || lock.current) return
    lock.current = true; setBusy(true); setError('')
    try {
      await request(`${resourcePath(kind, archiving.id)}/archive`, token, { method: 'PATCH', body: JSON.stringify({ is_archived: !archiving.is_archived }) })
      setMessage(archiving.is_archived ? 'Élément restauré.' : 'Élément archivé.'); setArchiving(null)
      await client.invalidateQueries()
    } catch (failure) { setError(failure instanceof Error ? failure.message : 'Opération impossible.') }
    finally { lock.current = false; setBusy(false) }
  }
  if (!isAdmin(session)) return null
  return <Box sx={{ mt: 2 }}>
    <Button startIcon={<TuneRounded />} onClick={() => { const show = () => { setError(''); setOpen(true) }; if (guard) guard.request(show); else show() }}>Gestion avancée</Button>
    <Dialog open={open && !archiving && !deleting} fullWidth maxWidth="lg" aria-labelledby="advanced-maintenance-title" onClose={() => { if (!busy) setOpen(false) }}>
      <DialogTitle id="advanced-maintenance-title">Gestion avancée</DialogTitle>
      <DialogContent>
        <Typography variant="body2" color="text.secondary" sx={{ mb: 2 }}>L’archivage conserve les informations. La suppression définitive est réservée aux super-administrateurs et aux éléments déjà archivés ; les dépendances sont vérifiées par le serveur.</Typography>
        <Box sx={{ display: 'flex', alignItems: 'center', gap: 2, flexWrap: 'wrap', mb: 2 }}>
          <TextField select size="small" label="Ressource" value={kind} disabled={busy} onChange={event => { setKind(event.target.value as Kind); setPage(0); setError('') }} sx={{ minWidth: 200 }}>{labels.map(([value, label]) => <MenuItem key={value} value={value}>{label}</MenuItem>)}</TextField>
          <ToggleButtonGroup exclusive size="small" value={archive} disabled={busy} aria-label="Dossier des ressources" onChange={(_, value: 'active' | 'archived' | null) => { if (value) { setArchive(value); setPage(0); setError('') } }}><ToggleButton value="active">Normal</ToggleButton><ToggleButton value="archived">Archivé</ToggleButton></ToggleButtonGroup>
        </Box>
        {kind === 'leave_request' && <Alert severity="info" sx={{ mb: 2 }}>Une demande en attente doit d’abord être traitée avant son archivage.</Alert>}
        {kind === 'annual_interview' && <Alert severity="info" sx={{ mb: 2 }}>La suppression d’un entretien requiert que son compte soit désactivé.</Alert>}
        {error && <Alert severity="error" sx={{ mb: 2 }}>{error}</Alert>}
        {list.isPending ? <CircularProgress size={26} /> : list.error ? <Alert severity="error">{list.error.message}</Alert> : <>
          <TableContainer sx={{ maxHeight: '55dvh' }}><Table size="small" stickyHeader aria-label="Ressources à archiver"><TableHead><TableRow><TableCell>Élément</TableCell><TableCell>Informations</TableCell><TableCell align="right" sx={{ width: 112 }}>Actions</TableCell></TableRow></TableHead><TableBody>
            {!list.data.items.length && <TableRow><TableCell colSpan={3}>Aucun élément dans ce dossier.</TableCell></TableRow>}
            {list.data.items.map(item => <TableRow key={item.id}><TableCell sx={{ overflowWrap: 'anywhere' }}>{item.label}<Typography variant="caption" sx={{ display: 'block' }} color="text.secondary">N° {item.id}</Typography></TableCell><TableCell sx={{ overflowWrap: 'anywhere' }}>{item.detail || '—'}</TableCell><TableCell align="right" sx={{ whiteSpace: 'nowrap' }}>
              <Tooltip title={item.is_archived ? 'Restaurer' : 'Archiver'}><span><IconButton aria-label={`${item.is_archived ? 'Restaurer' : 'Archiver'} ${item.label}`} disabled={busy} onClick={() => { setError(''); setArchiving({ ...item }) }}>{item.is_archived ? <UnarchiveOutlined fontSize="small" /> : <ArchiveOutlined fontSize="small" />}</IconButton></span></Tooltip>
              {isSuperAdmin(session) && item.is_archived && <DeleteAction label={`Supprimer définitivement ${item.label}`} disabled={busy} onClick={() => setDeleting({ ...item })} />}
            </TableCell></TableRow>)}
          </TableBody></Table></TableContainer>
        </>}
        <TablePagination component="div" count={list.data?.total ?? 0} page={list.data ? page : 0} rowsPerPage={25} rowsPerPageOptions={[25]} labelRowsPerPage="Par page" onPageChange={(_, next) => setPage(next)} labelDisplayedRows={({ from, to, count }) => `${from}–${to} sur ${count}`} />
      </DialogContent><DialogActions><Button disabled={busy} onClick={() => setOpen(false)}>Fermer</Button></DialogActions>
    </Dialog>
    {archiving && <Dialog open onClose={() => { if (!busy) setArchiving(null) }}><DialogTitle>{archiving.is_archived ? 'Restaurer cet élément ?' : 'Archiver cet élément ?'}</DialogTitle><DialogContent><Typography sx={{ overflowWrap: 'anywhere' }}>{archiving.label}</Typography><Alert severity="info" sx={{ mt: 2 }}>Les données restent conservées. Le serveur vérifie les conditions propres à cet élément.</Alert>{error && <Alert severity="error" sx={{ mt: 2 }}>{error}</Alert>}</DialogContent><DialogActions><Button disabled={busy} onClick={() => setArchiving(null)}>Annuler</Button><Button disabled={busy} onClick={() => void updateArchive()}>Confirmer</Button></DialogActions></Dialog>}
    {deleting && <PermanentDeleteDialog title="Supprimer définitivement cet élément ?" expected={deleting.id} description={`${deleting.label} sera supprimé, sous réserve des contrôles de dépendances.`} onClose={() => setDeleting(null)} onConfirm={confirmation => request(resourcePath(kind, deleting.id), token, { method: 'DELETE', body: JSON.stringify({ confirmation }) })} onDeleted={() => { setDeleting(null); setMessage('Élément supprimé.'); void client.invalidateQueries() }} />}
    <Snackbar open={!!message} autoHideDuration={5000} onClose={() => setMessage('')}><Alert severity="success" onClose={() => setMessage('')}>{message}</Alert></Snackbar>
  </Box>
}
