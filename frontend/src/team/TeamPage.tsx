import { useId, useState } from 'react'
import { Alert, Box, Button, CircularProgress, MenuItem, Paper, Tab, Table, TableBody, TableCell, TableContainer, TableHead, TableRow, Tabs, TextField, Typography } from '@mui/material'
import { useQuery } from '@tanstack/react-query'
import type { Session } from '../types'
import { teamApi } from './api'
import { OrganizationCards } from './OrganizationCards'
import { DepartmentBadges } from '../components/DepartmentBadge'
import { PersonProfile } from './PersonProfile'

export function TeamPage({ token, session, onAttendance }: { token: string; session: Session; onAttendance: () => void }) {
  const privileged = session.role_tags.some(role => ['Admin', 'Responsable'].includes(role))
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
  const rows = privileged ? (people.data?.users || []).filter(p => p.role_tags?.includes('Employe') || (!p.role_tags?.length && p.role === 'Employe')).map(p => [
    p.username, p.employee_type || 'salarie', p.manager_id || '', groupNames(p.group_ids || []), p.email,
  ]) : (team.data?.members || []).map(p => [p.name, groupNames([...p.group_ids, ...p.managed_group_ids]), groupNames(p.managed_group_ids)])
  const filtered = rows.filter(row => row.join(' ').toLocaleLowerCase('fr').includes(search.trim().toLocaleLowerCase('fr')))
  const pending = privileged ? people.isPending || groups.isPending : team.isPending
  const error = privileged ? people.error || groups.error : team.error
  const refresh = () => { void directory.refetch(); if (privileged) { void people.refetch(); void groups.refetch() } else void team.refetch() }
  return <Box className="workspace">
    <Box className="page-heading"><Box><Typography variant="overline" color="primary.main" sx={{ fontWeight: 700 }}>ORGANISATION</Typography><Typography variant="h4" component="h1">Mon équipe</Typography><Typography color="text.secondary" sx={{ mt: 1 }}>L’organisation de l’entreprise et votre équipe.</Typography></Box><Box sx={{ display: 'flex', gap: 1 }}><Button variant="outlined" onClick={() => setProfile(session.username)}>Mon profil mensuel</Button><Button variant="outlined" onClick={refresh}>Actualiser</Button></Box></Box>
    <Tabs value={sheet} onChange={(_, value: 'team' | 'organization') => setSheet(value)} variant="scrollable" scrollButtons="auto" aria-label="Vues de Mon équipe" sx={{ borderBottom: 1, borderColor: 'divider', mb: 2 }}>
      <Tab value="team" label={privileged ? 'Employés gérés' : 'Mon département'} id={`${tabsId}-team-tab`} aria-controls={`${tabsId}-team-panel`} />
      <Tab value="organization" label="Organigramme de l’entreprise" id={`${tabsId}-organization-tab`} aria-controls={`${tabsId}-organization-panel`} />
    </Tabs>
    <Box role="tabpanel" hidden={sheet !== 'team'} id={`${tabsId}-team-panel`} aria-labelledby={`${tabsId}-team-tab`}>
    <Paper sx={{ p: { xs: 2, md: 3 } }}>
      <Box sx={{ display: 'flex', flexWrap: 'wrap', alignItems: 'center', gap: 2, mb: 2 }}><Typography variant="h6" sx={{ flex: 1 }}>{privileged ? 'Employés gérés' : 'Mon département'}</Typography>
        {privileged && <TextField select label="Consulter un profil" value="" sx={{ minWidth: 200 }} onChange={event => setProfile(event.target.value)}>
          {(people.data?.users || []).filter(person => session.role_tags.includes('Admin') || person.username === session.username || person.manager_id === session.username).map(person => <MenuItem key={person.username} value={person.username}>{person.username}</MenuItem>)}
        </TextField>}
        <TextField label="Rechercher une personne" value={search} onChange={e => setSearch(e.target.value)} /></Box>
      {privileged && <Typography variant="body2" color="text.secondary" sx={{ mb: 2 }}>Les utilisateurs affichés appartiennent à votre périmètre de responsabilité. Les comptes et contrats sont gérés par un administrateur.</Typography>}
      {pending ? <CircularProgress size={24} /> : error ? <Alert severity="error">{error.message}<Button onClick={refresh}>Réessayer</Button></Alert> : !filtered.length ? <Typography color="text.secondary">{search ? 'Aucun résultat pour cette recherche.' : privileged ? 'Aucun Employé dans votre périmètre.' : 'Aucun membre dans votre département actif.'}</Typography> :
        <TableContainer><Table size="small" aria-label={privileged ? 'Employés gérés' : 'Mon département'}><TableHead><TableRow>{(privileged ? ['Utilisateur', 'Type', 'Responsable', 'Groupes', 'E-mail'] : ['Nom', 'Département', 'Responsable de']).map(header => <TableCell key={header}>{header}</TableCell>)}</TableRow></TableHead><TableBody>{filtered.map((row, i) => <TableRow key={i}>{row.map((value, j) => <TableCell key={j} sx={{ overflowWrap: 'anywhere' }}>{privileged && j === 0 ? <Button onClick={() => setProfile(value)} sx={{ textAlign: 'left', justifyContent: 'start', overflowWrap: 'anywhere' }}>{value}</Button> : (privileged ? j === 3 : j === 1) ? <DepartmentBadges names={value ? value.split(', ') : []} /> : value || '—'}</TableCell>)}</TableRow>)}</TableBody></Table></TableContainer>}
      {privileged && <Button sx={{ mt: 2 }} onClick={onAttendance}>Créer / remplir une fiche de présence</Button>}
    </Paper>
    </Box>
    <Box role="tabpanel" hidden={sheet !== 'organization'} id={`${tabsId}-organization-panel`} aria-labelledby={`${tabsId}-organization-tab`}>
      {directory.isPending ? <CircularProgress size={24} /> : directory.error ? <Alert severity="error">{directory.error.message}<Button onClick={() => void directory.refetch()}>Réessayer</Button></Alert> : directory.data && <OrganizationCards model={directory.data.organization} />}
    </Box>
    {profile && <PersonProfile key={profile} username={profile} token={token} viewer={session.username} onClose={() => setProfile(null)} />}
  </Box>
}
