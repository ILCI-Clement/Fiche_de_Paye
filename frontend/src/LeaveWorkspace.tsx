import { personDisplayName } from './personDisplayName'
import { isAdmin } from './rolePolicy'
import { useState } from 'react'
import { Alert, Avatar, Box, Button, CircularProgress, Divider, IconButton, Pagination, Paper, Snackbar, Tab, Tabs, Tooltip, Typography } from '@mui/material'
import AddRounded from '@mui/icons-material/AddRounded'
import CalendarMonthRounded from '@mui/icons-material/CalendarMonthRounded'
import CheckCircleOutlineRounded from '@mui/icons-material/CheckCircleOutlineRounded'
import RefreshRounded from '@mui/icons-material/RefreshRounded'
import ArrowForwardRounded from '@mui/icons-material/ArrowForwardRounded'
import AccessTimeRounded from '@mui/icons-material/AccessTimeRounded'
import { useQuery, useQueryClient } from '@tanstack/react-query'
import { api } from './api'
import type { Filter, Scope, Session } from './types'
import { initials, periodLabel, slotLabel, timeLabel } from './leaveFormat'
import { LeaveForm } from './components/LeaveForm'
import { LeaveDetail } from './components/LeaveDetail'
import { StatusChip } from './components/StatusChip'
import { alpha } from '@mui/material/styles'
import { WorkspaceToolbar } from './components/WorkspaceToolbar'


export function LeaveWorkspace({ token, session }: { token: string; session: Session }) {
  const client = useQueryClient()
  const [scope, setScope] = useState<Scope>('mine')
  const [filter, setFilter] = useState<Filter>('all')
  const [page, setPage] = useState(1)
  const [creating, setCreating] = useState(false)
  const [selected, setSelected] = useState<number | null>(null)
  const [message, setMessage] = useState('')
  const list = useQuery({ queryKey: ['leave-list', session.username, scope, filter, page],
    queryFn: ({ signal }) => api.list(token, scope, filter, page, signal), refetchInterval: 30000 })
  const reviews = useQuery({ queryKey: ['leave-list', session.username, 'review-count'],
    queryFn: ({ signal }) => api.list(token, 'team', 'pending', 1, signal), enabled: session.can_review, refetchInterval: 30000 })
  const pending = reviews.data?.counts.pending ?? 0
  function navigate(next: Scope) { setScope(next); setFilter('all'); setPage(1) }
  function saved(text: string) {
    setMessage(text); setCreating(false); setSelected(null); setPage(1)
    void client.invalidateQueries({ queryKey: ['leave-list'] })
    void client.invalidateQueries({ queryKey: ['leave-detail'] })
  }
  return <Box>
      <Box component="main" className="workspace">
        <WorkspaceToolbar>
          {session.can_review && <Tabs value={scope} onChange={(_, value: Scope) => navigate(value)} aria-label="Périmètre des demandes de congé" variant="scrollable" scrollButtons="auto" sx={{ minHeight: 36, '& .MuiTab-root': { minHeight: 36, px: 1.5, fontSize: 13 } }}><Tab value="mine" label="Mes demandes" /><Tab value="team" label={isAdmin(session) ? 'Toutes les demandes' : 'Mon équipe'} /></Tabs>}
          <Tooltip title="Actualiser les demandes"><span><IconButton aria-label="Actualiser les demandes" disabled={list.isFetching} onClick={() => { void list.refetch(); if (session.can_review) void reviews.refetch() }}><RefreshRounded fontSize="small" /></IconButton></span></Tooltip>
          <Button size="small" variant="contained" startIcon={<AddRounded />} onClick={() => setCreating(true)}>Nouvelle demande</Button>
        </WorkspaceToolbar>
        <Box className="stats-grid">
          {[{ label: 'En attente', value: list.data?.counts.pending, icon: <AccessTimeRounded />, color: 'warning' as const },
            { label: 'Approuvées', value: list.data?.counts.approved, icon: <CheckCircleOutlineRounded />, color: 'success' as const },
            { label: 'Total des demandes', value: list.data?.counts.total, icon: <CalendarMonthRounded />, color: 'secondary' as const }].map(stat =>
              <Paper key={stat.label} className="stat-card"><Box sx={theme => ({ bgcolor: alpha(theme.palette[stat.color].main, .14), color: theme.palette[stat.color].main, borderRadius: 3, p: 1.3, display: 'flex' })}>{stat.icon}</Box>
                <Box><Typography variant="body2" color="text.secondary">{stat.label}</Typography><Typography sx={{ fontWeight: 750, fontSize: 28 }}>{stat.value ?? '—'}</Typography></Box></Paper>)}
        </Box>
        {scope === 'team' && pending > 0 && <Alert severity="info" icon={<AccessTimeRounded />} sx={{ mb: 2 }}>
          {pending} demande{pending > 1 ? 's' : ''} attend{pending > 1 ? 'ent' : ''} votre décision.
        </Alert>}
        <Paper className="requests-panel">
          <Box className="list-heading"><Typography variant="h6">{scope === 'mine' ? 'Historique de mes demandes' : 'Demandes de congé'}</Typography></Box>
          <Tabs value={filter} onChange={(_, value: Filter) => { setFilter(value); setPage(1) }} variant="scrollable" scrollButtons="auto" sx={{ px: 2 }}>
            <Tab value="all" label="Toutes" /><Tab value="pending" label="En attente" /><Tab value="approved" label="Approuvées" /><Tab value="rejected" label="Refusées" />
          </Tabs><Divider />
          {list.isPending && <Box className="empty-state"><CircularProgress size={25} /><Typography color="text.secondary">Chargement des demandes…</Typography></Box>}
          {list.error && <Alert severity="error" sx={{ m: 3 }}>{list.error.message}</Alert>}
          {list.data?.items.length === 0 && <Box className="empty-state"><Box className="empty-icon"><CalendarMonthRounded sx={{ fontSize: 32 }} /></Box>
            <Typography variant="h6">Aucune demande {filter === 'pending' ? 'en attente' : 'à afficher'}</Typography>
            <Typography color="text.secondary">{scope === 'mine' && filter === 'all' ? 'Votre prochaine pause commence par une demande.' : 'Les demandes correspondant à ce filtre apparaîtront ici.'}</Typography>
            {scope === 'mine' && filter === 'all' && <Button variant="outlined" startIcon={<AddRounded />} onClick={() => setCreating(true)}>Créer ma première demande</Button>}
          </Box>}
          {Boolean(list.data?.items.length) && <>
            <Box className="request-columns"><span>COLLABORATEUR & PÉRIODE</span><span>RESPONSABLE</span><span>STATUT</span><span /></Box>
            {list.data?.items.map(item => <Box key={item.id} className="request-row">
              <Box sx={{ display: 'flex', gap: 1.5, minWidth: 0, alignItems: 'center' }}><Avatar sx={{ bgcolor: 'var(--surface-soft)', color: 'text.secondary', fontSize: 13 }}>{initials(personDisplayName({ display_name: item.employee_name }))}</Avatar>
                <Box sx={{ minWidth: 0 }}><Typography sx={{ fontWeight: 650, fontSize: 14 }}>{personDisplayName({ display_name: item.employee_name })}</Typography><Typography sx={{ fontSize: 13, mt: 0.4 }}>{periodLabel(item)}</Typography><Typography variant="caption" color="text.secondary">{slotLabel(item)}</Typography></Box></Box>
              <Box className="request-manager"><Typography variant="body2">{personDisplayName({ display_name: item.approver_name })}</Typography><Typography variant="caption" color="text.secondary">Envoyée le {timeLabel(item.created_at)}</Typography></Box>
              <Box><StatusChip status={item.status} /></Box>
              <Button className="detail-button" size="small" color="inherit" endIcon={<ArrowForwardRounded fontSize="small" />} onClick={() => setSelected(item.id)}>{item.can_review ? 'Traiter' : 'Détails'}</Button>
            </Box>)}
          </>}
          {list.data && (list.data.total > 12 || page > 1) && <Box sx={{ p: 2, display: 'flex', justifyContent: 'center' }}><Pagination page={page} count={Math.max(1, Math.ceil(list.data.total / 12))} onChange={(_, value) => setPage(value)} /></Box>}
        </Paper>
      </Box>
    {creating && <LeaveForm token={token} session={session} onClose={() => setCreating(false)} onSaved={() => saved('Votre demande a été envoyée.')} />}
    {selected !== null && <LeaveDetail token={token} id={selected} onClose={() => setSelected(null)} onChanged={saved} />}
    <Snackbar open={Boolean(message)} autoHideDuration={4500} onClose={() => setMessage('')}><Alert severity="success" variant="filled" onClose={() => setMessage('')}>{message}</Alert></Snackbar>
  </Box>
}
