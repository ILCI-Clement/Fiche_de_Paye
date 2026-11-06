import { useState } from 'react'
import { useWorkNavigation } from '../workspaceNavigation'
import { Alert, Box, Button, CircularProgress, IconButton, Paper, Snackbar, TablePagination, Tooltip, Typography } from '@mui/material'
import DownloadOutlined from '@mui/icons-material/DownloadOutlined'
import HistoryRounded from '@mui/icons-material/HistoryRounded'
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
import { PersonnelSnapshots } from './PersonnelSnapshots'

export function PersonnelPage({ token, session }: { token: string; session: Session }) {
  if (!isAdmin(session)) return <Box className="workspace"><Alert severity="error">La liste des utilisateurs est réservée aux administrateurs.</Alert></Box>
  return <RegistryViewerContext value={session.username}><PersonnelWorkspace token={token} viewer={session.username} canDelete={isSuperAdmin(session)} /></RegistryViewerContext>
}

function PersonnelWorkspace({ token, viewer, canDelete }: { token: string; viewer: string; canDelete: boolean }) {
  const navigation = useWorkNavigation()
  const [filters, setFilters] = useState<PersonnelFilters>({ q: '', personnel_type: '', archive: 'active', status: 'all', page: 1, page_size: 25 })
  const [search, setSearch] = useState('')
  const [creating, setCreating] = useState(false)
  const [exporting, setExporting] = useState(false)
  const [snapshots, setSnapshots] = useState(false)
  const [selected, setSelected] = useState<number | null>(null)
  const [appliedWork, setAppliedWork] = useState<number | null>(null)
  if (navigation?.request?.page === 'personnel' && appliedWork !== navigation.request.sequence) {
    setAppliedWork(navigation.request.sequence)
    setSelected(navigation.request.employeeId || null)
    setSnapshots(navigation.request.snapshots === true)
  }
  const [message, setMessage] = useState('')
  const list = useQuery({ queryKey: ['personnel-register', viewer, 'list', filters], queryFn: ({ signal }) => personnelApi.list(token, filters, signal), retry: false })
  // An archive on the last page may make that page empty while earlier pages remain.
  const lastPage = list.data ? Math.max(1, Math.ceil(list.data.total / filters.page_size)) : filters.page
  if (filters.page > lastPage) setFilters({ ...filters, page: lastPage })
  const change = (update: Partial<PersonnelFilters>) => setFilters(current => ({ ...current, ...update, page: 1 }))
  const selectedFilters = !!filters.q || !!filters.personnel_type || filters.status !== 'all' || filters.archive !== 'active'
  const tableActions = <Box role="group" aria-label="Exports et archives" sx={{ display: 'flex', alignItems: 'center', gap: .5, px: 1, flexShrink: 0 }}>
    <Tooltip title="Exporter les utilisateurs correspondant aux filtres, toutes pages incluses"><span>
      <IconButton size="small" aria-label="Exporter les utilisateurs" color="primary" disabled={list.isFetching || !list.data?.total} onClick={() => setExporting(true)}><DownloadOutlined fontSize="small" /></IconButton>
    </span></Tooltip>
    <Tooltip title="Historique des archives"><IconButton size="small" aria-label="Historique des archives" color="primary" onClick={() => setSnapshots(true)}><HistoryRounded fontSize="small" /></IconButton></Tooltip>
  </Box>
  return <Box className="workspace personnel-workspace" sx={{
    '&.workspace': { maxWidth: 'none', m: 0, p: { xs: 1, sm: '12px 16px' }, height: 'max(0px, calc(100dvh - var(--workspace-toolbar-height, 80px)))', minHeight: 0, display: 'flex', flexDirection: 'column' },
    '& .MuiButton-root, & .MuiTypography-h6': { fontWeight: 500 }, '& .MuiTableCell-root': { py: .75, px: 1.25, fontSize: 13 }, '& .MuiTableCell-head': { fontWeight: 500, py: 1 },
  }}>
    <WorkspaceToolbar><PersonnelToolbar filters={filters} search={search} busy={list.isFetching}
      onSearchChange={setSearch} onSearch={() => change({ q: search.trim() })} onFilter={change}
      onReset={() => { setSearch(''); setFilters(current => ({ ...current, q: '', personnel_type: '', status: 'all', archive: 'active', page: 1 })) }}
      onRegister={() => setCreating(true)} onRefresh={() => void list.refetch()} /></WorkspaceToolbar>
    <Paper sx={{ border: 1, borderColor: 'divider', overflow: 'hidden', display: 'flex', flexDirection: 'column', flex: 1, minHeight: 0 }}>
      {!list.data?.employees.length && <Box sx={{ display: 'flex', justifyContent: 'flex-end', minHeight: 40, alignItems: 'center', borderBottom: 1, borderColor: 'divider', flexShrink: 0 }}>{tableActions}</Box>}
      {list.isPending ? <Box sx={{ p: 5, textAlign: 'center' }}><CircularProgress size={28} aria-label="Chargement du registre" /></Box> : list.error ? <Alert severity="error" sx={{ m: 2 }}>{list.error.message}<Button onClick={() => void list.refetch()}>Réessayer</Button></Alert> : !list.data.employees.length ? <Box sx={{ px: 3, py: 6, textAlign: 'center' }}>
        <Typography sx={{ fontWeight: 500, mb: 1 }}>{selectedFilters ? 'Aucune personne ne correspond à ces filtres.' : 'Le registre ne contient pas encore de personnel.'}</Typography>
        <Typography variant="body2" color="text.secondary">{selectedFilters ? 'Modifiez les filtres ou consultez les archives.' : 'Enregistrez une personne avec ses informations et sa date d’entrée.'}</Typography>
        {filters.page > 1 && <Button sx={{ mt: 2 }} onClick={() => setFilters(current => ({ ...current, page: 1 }))}>Revenir à la première page</Button>}
        {!selectedFilters && <Button sx={{ mt: 2 }} onClick={() => setCreating(true)}>Enregistrer une première personne</Button>}
      </Box> : <PersonnelTable employees={list.data.employees} token={token} headerActions={tableActions}
        active={!list.isFetching && selected === null && !creating && !exporting}
        resetKey={JSON.stringify([filters, selected, creating, exporting])} onOpen={setSelected} />}
      <Box sx={{ mt: 'auto', flexShrink: 0, display: 'flex', justifyContent: 'flex-end', borderTop: 1, borderColor: 'divider', px: 1.5, py: .5 }}>
        <TablePagination component="div" count={list.data?.total ?? 0} page={list.data ? filters.page - 1 : 0} rowsPerPage={filters.page_size} rowsPerPageOptions={[25, 50, 100]}
          disabled={list.isFetching || !list.data} onPageChange={(_, page) => setFilters(current => ({ ...current, page: page + 1 }))} onRowsPerPageChange={event => change({ page_size: Number(event.target.value) })}
          labelRowsPerPage="Lignes" labelDisplayedRows={({ from, to, count }) => list.data ? `${from}–${to} sur ${count}` : list.isFetching ? 'Chargement…' : '—'}
          getItemAriaLabel={type => type === 'next' ? 'Page suivante' : type === 'previous' ? 'Page précédente' : type === 'first' ? 'Première page' : 'Dernière page'}
          sx={{ ml: { xs: 0, md: 'auto' }, maxWidth: '100%', border: 0,
            '& .MuiTablePagination-toolbar': { minHeight: 40, p: 0, pl: 1, gap: .5, flexWrap: 'wrap' },
            '& .MuiTablePagination-spacer': { display: 'none' }, '& .MuiTablePagination-selectLabel': { m: 0, fontSize: 12 },
            '& .MuiTablePagination-select': { py: .5, pl: 1, pr: 2.5, fontSize: 12 },
            '& .MuiTablePagination-input': { mx: .5 }, '& .MuiTablePagination-displayedRows': { m: 0, fontSize: 12, fontVariantNumeric: 'tabular-nums' },
            '& .MuiTablePagination-actions': { ml: .5 },
          }} />
      </Box>
    </Paper>
    {creating && <PersonnelEditor token={token} allowSuperAdmin={canDelete} onClose={() => setCreating(false)} onSaved={setMessage} />}
    {snapshots && <PersonnelSnapshots token={token} viewer={viewer} onClose={() => setSnapshots(false)} />}
    {exporting && <PersonnelExport token={token} filters={filters} total={list.data?.total ?? 0} canExportExcel={canDelete} onClose={() => setExporting(false)} onSuccess={() => setMessage('Export généré. Téléchargement lancé.')} />}
    {selected !== null && <PersonnelDetail key={selected} token={token} employeeId={selected} completion={navigation?.request?.page === 'personnel' && navigation.request.employeeId === selected ? navigation.request : undefined} canDelete={canDelete} onClose={() => setSelected(null)} onSaved={setMessage} />}
    <Snackbar open={!!message} autoHideDuration={6000} onClose={() => setMessage('')}><Alert severity="success" variant="filled" onClose={() => setMessage('')}>{message}</Alert></Snackbar>
  </Box>
}
