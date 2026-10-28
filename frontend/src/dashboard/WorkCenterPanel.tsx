import { useState } from 'react'
import { Alert, Box, Button, Chip, CircularProgress, IconButton, Pagination, Paper, TextField, Tooltip, Typography } from '@mui/material'
import RefreshRounded from '@mui/icons-material/RefreshRounded'
import HelpOutlineRounded from '@mui/icons-material/HelpOutlineRounded'
import { PeriodToolbar } from '../attendance/PeriodToolbar'
import { useWorkNavigation } from '../workspaceNavigation'
import { useWorkCenter } from './useWorkCenter'
import { ReadinessStatus } from './ReadinessStatus'
import type { TaskKind } from './workCenterApi'

const kinds: { value: TaskKind; label: string }[] = [
  { value: 'leave', label: 'Congés à valider' }, { value: 'signature', label: 'Signatures à suivre' },
  { value: 'archive', label: 'PDF à récupérer' }, { value: 'readiness', label: 'Dossiers à vérifier' },
]
const priority = { leave: 0, archive: 1, signature: 2, readiness: 3 }

export function WorkCenterPanel({ token, viewer }: { token: string; viewer: string }) {
  const [period, setPeriod] = useState(() => ({ month: new Date().getMonth() + 1, year: new Date().getFullYear() }))
  const [kind, setKind] = useState<TaskKind | 'all'>('all')
  const [page, setPage] = useState(1)
  const [search, setSearch] = useState('')
  const list = useWorkCenter(token, viewer, period.month, period.year)
  const navigation = useWorkNavigation()
  const tasks = (list.data?.tasks || []).filter(task => (kind === 'all' || task.kind === kind) && `${task.name} ${task.label}`.toLocaleLowerCase('fr').includes(search.trim().toLocaleLowerCase('fr')))
    .sort((a, b) => priority[a.kind] - priority[b.kind])
  const shownPage = Math.min(page, Math.max(1, Math.ceil(tasks.length / 8)))
  const people = new Map(list.data?.people.map(person => [person.username, person]))
  return <Paper component="section" aria-label="Centre de tâches" variant="outlined" sx={{ p: { xs: 1.5, sm: 2 }, mb: 1.5, minWidth: 0 }}>
    <Box sx={{ display: 'flex', alignItems: 'center', gap: 1, mb: 1.5 }}><Typography component="h2" sx={{ fontSize: 17, fontWeight: 600 }}>À traiter</Typography>
      <Tooltip title="Congés et signatures : toutes périodes, dans votre périmètre autorisé. Dossiers : contrats du mois choisi ; le registre reste facultatif pour créer une fiche." arrow><IconButton size="small" aria-label="À propos des tâches"><HelpOutlineRounded fontSize="small" /></IconButton></Tooltip>
      <Tooltip title="Actualiser les tâches"><span><IconButton size="small" aria-label="Actualiser les tâches" disabled={list.isFetching} onClick={() => void list.refetch()}><RefreshRounded fontSize="small" /></IconButton></span></Tooltip>
    </Box>
    <Box sx={{ display: 'grid', gridTemplateColumns: { xs: 'repeat(2, minmax(0, 1fr))', md: 'repeat(4, minmax(0, 1fr))' }, gap: 1, mb: 1.5 }}>
      {kinds.map(item => <Button key={item.value} variant={kind === item.value ? 'contained' : 'outlined'} color={kind === item.value ? 'primary' : 'inherit'} onClick={() => { setKind(kind === item.value ? 'all' : item.value); setPage(1) }} aria-pressed={kind === item.value} sx={{ justifyContent: 'space-between', textAlign: 'left', px: 1.5, py: 1, gap: 1, minWidth: 0 }}>
        <Typography component="span" sx={{ fontSize: 13 }}>{item.label}</Typography><Typography component="span" sx={{ fontSize: 22, fontWeight: 600 }}>{list.data ? list.data.counts[item.value] : '—'}</Typography>
      </Button>)}
    </Box>
    <Box sx={{ display: 'flex', flexWrap: 'wrap', gap: 1, alignItems: 'center', mb: 1 }}>
      <TextField size="small" placeholder="Rechercher une tâche" value={search} onChange={event => { setSearch(event.target.value); setPage(1) }} slotProps={{ htmlInput: { 'aria-label': 'Rechercher une tâche' } }} sx={{ flex: '1 1 180px', minWidth: 0 }} />
      <PeriodToolbar label="Mois de vérification des dossiers" month={period.month} year={period.year} disabled={list.isFetching} onChange={(month, year) => { setPeriod({ month, year }); setPage(1) }} />
    </Box>
    {list.isPending && <CircularProgress size={22} aria-label="Chargement des tâches" />}
    {list.error && <Alert severity="error">Impossible de vérifier les tâches. {list.error.message}<Button onClick={() => void list.refetch()}>Réessayer</Button></Alert>}
    {list.data && !list.error && <>
      {!tasks.length && <Typography color="text.secondary" variant="body2" sx={{ py: 2 }}>{search || kind !== 'all' ? 'Aucune tâche pour ce filtre.' : 'Aucune tâche à traiter dans votre périmètre.'}</Typography>}
      {tasks.slice((shownPage - 1) * 8, shownPage * 8).map(task => {
        const person = task.id.startsWith('person:') ? people.get(task.id.slice(7)) : undefined
        return <Box key={task.id} sx={{ display: 'flex', alignItems: 'center', gap: 1.5, py: 1, borderTop: 1, borderColor: 'divider', minWidth: 0 }}>
          <Box sx={{ flex: 1, minWidth: 0 }}><Box sx={{ display: 'flex', gap: 1, alignItems: 'center', flexWrap: 'wrap' }}><Typography variant="body2" sx={{ fontWeight: 500, overflowWrap: 'anywhere' }}>{task.name}</Typography><Chip size="small" variant="outlined" label={kinds.find(item => item.value === task.kind)?.label} sx={{ fontSize: 11, height: 22 }} /></Box>
          <Typography variant="caption" color="text.secondary" noWrap title={task.label} sx={{ display: 'block', mt: .25 }}>{task.label}</Typography></Box>
          {person ? <ReadinessStatus person={person} /> : task.destination && <Button size="small" sx={{ flexShrink: 0 }} onClick={() => navigation?.navigate(task.destination!)}>{task.kind === 'leave' ? 'Traiter' : 'Ouvrir'}</Button>}
        </Box>
      })}
      {tasks.length > 8 && <Box sx={{ display: 'flex', justifyContent: 'center', pt: 1 }}><Pagination size="small" page={shownPage} count={Math.ceil(tasks.length / 8)} onChange={(_, value) => setPage(value)} /></Box>}
      {kind === 'readiness' && <Box sx={{ mt: 1 }}><Typography variant="caption" color="text.secondary">{list.data.people.filter(person => !person.issues.length).length} dossier(s) sans point à vérifier. Le détail est aussi disponible dans Mon équipe et Administration.</Typography></Box>}
    </>}
  </Paper>
}
