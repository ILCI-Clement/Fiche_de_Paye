import { personDisplayName } from '../personDisplayName'
import { canManageTeam, isAdmin } from '../rolePolicy'
import { useId, useState } from 'react'
import { Alert, Box, Button, CircularProgress, IconButton, MenuItem, Paper, Tab, Table, TableBody, TableCell, TableContainer, TableHead, TableRow, Tabs, TextField, Tooltip, Typography } from '@mui/material'
import PersonOutlineRounded from '@mui/icons-material/PersonOutlineRounded'
import RefreshRounded from '@mui/icons-material/RefreshRounded'
import CalendarMonthOutlined from '@mui/icons-material/CalendarMonthOutlined'
import { useQuery } from '@tanstack/react-query'
import type { Session } from '../types'
import { teamApi } from './api'
import { OrganizationCards } from './OrganizationCards'
import { DepartmentBadges } from '../components/DepartmentBadge'
import { PersonProfile } from './PersonProfile'
import { WorkspaceToolbar } from '../components/WorkspaceToolbar'

export function TeamPage({ token, session, onAttendance }: { token: string; session: Session; onAttendance: () => void }) {
  const privileged = canManageTeam(session)
  const [sheet, setSheet] = useState<'team' | 'organization'>('team')
  const tabsId = useId()
  const [search, setSearch] = useState('')
  const [profile, setProfile] = useState<string | null>(null)
  const directory = useQuery({ queryKey: ['organization', session.username], queryFn: ({ signal }) => teamApi.directory(token, signal), retry: false })
  const team = useQuery({ queryKey: ['my-team', session.username], queryFn: ({ signal }) => teamApi.mine(token, signal), enabled: !privileged, retry: false })
  const people = useQuery({ queryKey: ['team-personnel', session.username], queryFn: ({ signal }) => teamApi.people(token, signal), enabled: privileged, retry: false })
  const groups = useQuery({ queryKey: ['team-groups', session.username], queryFn: ({ signal }) => teamApi.groups(token, signal), enabled: privileged, retry: false })
  const names = new Map((privileged ? groups.data?.groups : team.data?.groups)?.filter(g => g.is_active === undefined || Boolean(g.is_active)).map(g => [g.id, g.name]))
  const groupNames = (ids: number[]) => [...new Set(ids)].map(id => names.get(id) || String(id)).join(', ')
  const rows = privileged ? (people.data?.users || []).filter(p => p.role_tags?.includes('Employe') || (!p.role_tags?.length && p.role === 'Employe')).map(p => ({ id: p.username, cells: [
    personDisplayName(p), p.employee_type || 'salarie', p.manager_id ? personDisplayName({ display_name: p.manager_display_name }) : '', groupNames(p.group_ids || []), p.email,
  ] })) : (team.data?.members || []).map((p, index) => ({ id: String(index), cells: [personDisplayName({ full_name: p.full_name, display_name: p.display_name || p.name }), groupNames([...p.group_ids, ...p.managed_group_ids]), groupNames(p.managed_group_ids)] }))
  const filtered = rows.filter(row => row.cells.join(' ').toLocaleLowerCase('fr').includes(search.trim().toLocaleLowerCase('fr')))
  const pending = privileged ? people.isPending || groups.isPending : team.isPending
  const error = privileged ? people.error || groups.error : team.error
  const refresh = () => { void directory.refetch(); if (privileged) { void people.refetch(); void groups.refetch() } else void team.refetch() }
  return <Box className="workspace">
    <WorkspaceToolbar>
      <Box className="workspace-action-bar" sx={{ width: '100%' }}>
        {sheet === 'team' && <>
          <TextField size="small" placeholder="Rechercher une personne" value={search} onChange={event => setSearch(event.target.value)} slotProps={{ htmlInput: { 'aria-label': 'Rechercher une personne' } }} sx={{ width: { xs: '100%', sm: 230 }, maxWidth: '100%' }} />
          {privileged && <TextField select size="small" label="Consulter un profil" value="" sx={{ minWidth: 'min(100%, 220px)', maxWidth: 360, flex: { xs: '1 1 100%', sm: '0 1 260px' } }} slotProps={{ select: { MenuProps: { slotProps: { paper: { sx: { width: 640, maxWidth: 'calc(100vw - 32px)' } } } } } }} onChange={event => setProfile(event.target.value)}>
            {(people.data?.users || []).filter(person => isAdmin(session) || person.username === session.username || person.manager_id === session.username).map(person => <MenuItem key={person.username} value={person.username} sx={{ display: 'grid', gridTemplateColumns: 'minmax(0, 1fr) minmax(0, 1.5fr)', columnGap: 2, alignItems: 'center', py: 1 }}>
              <Typography component="span" title={personDisplayName(person)} sx={{ minWidth: 0, fontSize: 14, fontWeight: 500, whiteSpace: 'normal', overflowWrap: 'anywhere' }}>{personDisplayName(person)}</Typography>
              <Typography component="span" title={person.email || ''} sx={{ minWidth: 0, fontSize: 13, color: 'text.secondary', whiteSpace: 'normal', overflowWrap: 'anywhere' }}>{person.email || '—'}</Typography>
            </MenuItem>)}
          </TextField>}
        </>}
        <Tooltip title="Consulter mon profil mensuel"><Button variant="outlined" startIcon={<PersonOutlineRounded />} onClick={() => setProfile(session.username)}>Mon profil</Button></Tooltip>
        {privileged && <Tooltip title="Créer / remplir une fiche de présence"><IconButton aria-label="Créer / remplir une fiche de présence" onClick={onAttendance}><CalendarMonthOutlined /></IconButton></Tooltip>}
        <Tooltip title="Actualiser"><IconButton aria-label="Actualiser" onClick={refresh}><RefreshRounded /></IconButton></Tooltip>
      </Box>
    </WorkspaceToolbar>
    <Tabs value={sheet} onChange={(_, value: 'team' | 'organization') => setSheet(value)} variant="scrollable" scrollButtons="auto" aria-label="Vues de Mon équipe" sx={{ borderBottom: 1, borderColor: 'divider', mb: 2 }}>
      <Tab value="team" label={privileged ? 'Employés gérés' : 'Mon département'} id={`${tabsId}-team-tab`} aria-controls={`${tabsId}-team-panel`} />
      <Tab value="organization" label="Organigramme de l’entreprise" id={`${tabsId}-organization-tab`} aria-controls={`${tabsId}-organization-panel`} />
    </Tabs>
    <Box role="tabpanel" hidden={sheet !== 'team'} id={`${tabsId}-team-panel`} aria-labelledby={`${tabsId}-team-tab`}>
    <Paper sx={{ p: { xs: 1, sm: 2 } }}>
      {pending ? <CircularProgress size={24} /> : error ? <Alert severity="error">{error.message}<Button onClick={refresh}>Réessayer</Button></Alert> : !filtered.length ? <Typography color="text.secondary">{search ? 'Aucun résultat pour cette recherche.' : privileged ? 'Aucun Employé dans votre périmètre.' : 'Aucun membre dans votre département actif.'}</Typography> :
        <TableContainer><Table size="small" aria-label={privileged ? 'Employés gérés' : 'Mon département'}><TableHead><TableRow>{(privileged ? ['Nom et prénoms', 'Type', 'Responsable', 'Groupes', 'E-mail'] : ['Nom', 'Département', 'Responsable de']).map(header => <TableCell key={header}>{header}</TableCell>)}</TableRow></TableHead><TableBody>{filtered.map(row => <TableRow key={row.id}>{row.cells.map((value, j) => <TableCell key={j} sx={{ overflowWrap: 'anywhere' }}>{privileged && j === 0 ? <Button onClick={() => setProfile(row.id)} sx={{ textAlign: 'left', justifyContent: 'start', overflowWrap: 'anywhere' }}>{value}</Button> : (privileged ? j === 3 : j === 1) ? <DepartmentBadges names={value ? value.split(', ') : []} /> : value || '—'}</TableCell>)}</TableRow>)}</TableBody></Table></TableContainer>}
    </Paper>
    </Box>
    <Box role="tabpanel" hidden={sheet !== 'organization'} id={`${tabsId}-organization-panel`} aria-labelledby={`${tabsId}-organization-tab`}>
      {directory.isPending ? <CircularProgress size={24} /> : directory.error ? <Alert severity="error">{directory.error.message}<Button onClick={() => void directory.refetch()}>Réessayer</Button></Alert> : directory.data && <OrganizationCards model={directory.data.organization} />}
    </Box>
    {profile && <PersonProfile key={profile} username={profile} token={token} viewer={session.username} onClose={() => setProfile(null)} />}
  </Box>
}
