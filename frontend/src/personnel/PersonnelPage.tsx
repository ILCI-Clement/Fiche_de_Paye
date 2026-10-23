import { useState } from 'react'
import { Alert, Box, Button, CircularProgress, Paper, Snackbar, Typography } from '@mui/material'
import { useQuery } from '@tanstack/react-query'
import type { Session } from '../types'
import { personnelApi } from './api'
import type { PersonnelFilters } from './api'
import { PersonnelEditor } from './PersonnelEditor'
import { PersonnelDetail } from './PersonnelDetail'
import { RegistryViewerContext } from './context'
import { WorkspaceToolbar } from '../components/WorkspaceToolbar'
import { PersonnelToolbar } from './PersonnelToolbar'
import { PersonnelExport } from './PersonnelExport'
import { PersonnelTable } from './PersonnelTable'
import { isAdmin, isSuperAdmin } from '../rolePolicy'

export function PersonnelPage({ token, session }: { token: string; session: Session }) {
  if (!isAdmin(session)) return <Box className="workspace"><Alert severity="error">Le registre du personnel est réservé aux administrateurs.</Alert></Box>
  return <RegistryViewerContext value={session.username}><PersonnelWorkspace token={token} viewer={session.username} canDelete={isSuperAdmin(session)} /></RegistryViewerContext>
}

function PersonnelWorkspace({ token, viewer, canDelete }: { token: string; viewer: string; canDelete: boolean }) {
  const [filters, setFilters] = useState<PersonnelFilters>({ q: '', personnel_type: '', archive: 'active', status: 'all', page: 1, page_size: 25 })
  const [search, setSearch] = useState('')
  const [creating, setCreating] = useState(false)
  const [exporting, setExporting] = useState(false)
  const [selected, setSelected] = useState<number | null>(null)
  const [message, setMessage] = useState('')
  const list = useQuery({ queryKey: ['personnel-register', viewer, 'list', filters], queryFn: ({ signal }) => personnelApi.list(token, filters, signal), retry: false })
  // An archive on the last page may make that page empty while earlier pages remain.
  const lastPage = list.data ? Math.max(1, Math.ceil(list.data.total / filters.page_size)) : filters.page
  if (filters.page > lastPage) setFilters({ ...filters, page: lastPage })
  const change = (update: Partial<PersonnelFilters>) => setFilters(current => ({ ...current, ...update, page: 1 }))
  const selectedFilters = !!filters.q || !!filters.personnel_type || filters.status !== 'all' || filters.archive !== 'active'
  return <Box className="workspace personnel-workspace" sx={{
    '&.workspace': { maxWidth: 'none', m: 0, p: { xs: 1, sm: '12px 16px' }, height: 'max(0px, calc(100dvh - var(--workspace-toolbar-height, 80px)))', minHeight: 0, display: 'flex', flexDirection: 'column' },
    '& .MuiButton-root, & .MuiTypography-h6': { fontWeight: 500 }, '& .MuiTableCell-root': { py: .75, px: 1.25, fontSize: 13 }, '& .MuiTableCell-head': { fontWeight: 500, py: 1 },
  }}>
    <WorkspaceToolbar><PersonnelToolbar filters={filters} search={search} total={list.data?.total ?? 0} busy={list.isFetching} hasData={!!list.data}
      onSearchChange={setSearch} onSearch={() => change({ q: search.trim() })} onFilter={change}
      onReset={() => { setSearch(''); setFilters(current => ({ ...current, q: '', personnel_type: '', status: 'all', archive: 'active', page: 1 })) }}
      onRegister={() => setCreating(true)} onExport={() => setExporting(true)} onRefresh={() => void list.refetch()} onPage={page => setFilters(current => ({ ...current, page }))} /></WorkspaceToolbar>
    <Paper sx={{ border: 1, borderColor: 'divider', overflow: 'hidden', display: 'flex', flexDirection: 'column', flex: 1, minHeight: 0 }}>
      {list.isPending ? <Box sx={{ p: 5, textAlign: 'center' }}><CircularProgress size={28} aria-label="Chargement du registre" /></Box> : list.error ? <Alert severity="error" sx={{ m: 2 }}>{list.error.message}<Button onClick={() => void list.refetch()}>Réessayer</Button></Alert> : !list.data.employees.length ? <Box sx={{ px: 3, py: 6, textAlign: 'center' }}>
        <Typography sx={{ fontWeight: 500, mb: 1 }}>{selectedFilters ? 'Aucune personne ne correspond à ces filtres.' : 'Le registre ne contient pas encore de personnel.'}</Typography>
        <Typography variant="body2" color="text.secondary">{selectedFilters ? 'Modifiez les filtres ou consultez les archives.' : 'Enregistrez une personne avec ses informations et sa date d’entrée.'}</Typography>
        {filters.page > 1 && <Button sx={{ mt: 2 }} onClick={() => setFilters(current => ({ ...current, page: 1 }))}>Revenir à la première page</Button>}
        {!selectedFilters && <Button sx={{ mt: 2 }} onClick={() => setCreating(true)}>Enregistrer une première personne</Button>}
      </Box> : <PersonnelTable employees={list.data.employees} token={token}
        active={!list.isFetching && selected === null && !creating && !exporting}
        resetKey={JSON.stringify([filters, selected, creating, exporting])} onOpen={setSelected} />}
    </Paper>
    {creating && <PersonnelEditor token={token} onClose={() => setCreating(false)} onSaved={setMessage} />}
    {exporting && <PersonnelExport token={token} filters={filters} total={list.data?.total ?? 0} onClose={() => setExporting(false)} onSuccess={() => setMessage('Export Excel généré. Téléchargement lancé.')} />}
    {selected !== null && <PersonnelDetail key={selected} token={token} employeeId={selected} canDelete={canDelete} onClose={() => setSelected(null)} onSaved={setMessage} />}
    <Snackbar open={!!message} autoHideDuration={6000} onClose={() => setMessage('')}><Alert severity="success" variant="filled" onClose={() => setMessage('')}>{message}</Alert></Snackbar>
  </Box>
}
