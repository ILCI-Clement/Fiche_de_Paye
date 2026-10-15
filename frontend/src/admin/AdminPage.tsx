import { useCapabilities } from '../capabilities'
import { CreateAccount } from './CreateAccount'
import { useDraftGuard, useWorkspaceGuard } from '../draftGuard'
import { useRef, useState } from 'react'
import { Accordion, AccordionDetails, AccordionSummary, Alert, Box, Button, CircularProgress, MenuItem, Paper, Snackbar, TextField, Typography } from '@mui/material'
import ExpandMoreRounded from '@mui/icons-material/ExpandMoreRounded'
import PersonAddOutlined from '@mui/icons-material/PersonAddOutlined'
import { useQuery, useQueryClient } from '@tanstack/react-query'
import type { Session } from '../types'
import { adminApi } from './api'
import { GroupEditor } from './GroupEditor'
import { UserEditor } from './UserEditor'

export function AdminPage({ token, session }: { token: string; session: Session }) {
  if (!session.role_tags.includes('Admin')) return <Box className="workspace"><Alert severity="error">Cette page est réservée aux administrateurs.</Alert></Box>
  return <AdminWorkspace token={token} session={session} />
}

function AdminWorkspace({ token, session }: { token: string; session: Session }) {
  const client = useQueryClient()
  const capabilities = useCapabilities()
  const users = useQuery({ queryKey: ['admin-users', session.username], queryFn: ({ signal }) => adminApi.users(token, signal), retry: false })
  const groups = useQuery({ queryKey: ['admin-groups', session.username], queryFn: ({ signal }) => adminApi.groups(token, signal), retry: false })
  const [selected, setSelected] = useState('')
  const [search, setSearch] = useState('')
  const [creating, setCreating] = useState(false)
  const [busy, setBusy] = useState(false)
  const lock = useRef(false)
  const [error, setError] = useState('')
  const [message, setMessage] = useState('')
  const guard = useWorkspaceGuard()
  useDraftGuard(false, busy, 'Administration')
  async function run(action: () => Promise<unknown>, success: string) {
    if (lock.current) return false
    lock.current = true; setBusy(true); setError('')
    try { await action(); setMessage(success); await client.invalidateQueries(); return true }
    catch (e) { setError(e instanceof Error ? e.message : 'Une erreur est survenue.'); return false }
    finally { lock.current = false; setBusy(false) }
  }
  const current = users.data?.users.find(p => p.username === selected) || users.data?.users[0]
  const choices = (users.data?.users || []).filter(p => p.username === current?.username || `${p.username} ${p.email}`.toLocaleLowerCase('fr').includes(search.toLocaleLowerCase('fr')))
  return <Box className="workspace">
    <Box className="page-heading"><Box><Typography variant="overline" color="primary.main" sx={{ fontWeight: 700 }}>ADMINISTRATION</Typography><Typography variant="h4" component="h1">Comptes et départements</Typography></Box>
      <Box sx={{ display: 'flex', flexWrap: 'wrap', gap: 1 }}>
        <Button variant="outlined" startIcon={<PersonAddOutlined />} disabled={busy || !users.data || !groups.data || !!users.error || !!groups.error} onClick={() => { setError(''); setCreating(true) }}>Créer un compte</Button>
        <Button disabled={busy} onClick={() => { guard?.request(() => { void users.refetch(); void groups.refetch() }) }}>Actualiser</Button>
      </Box>
    </Box>
    {!capabilities.canCreateAccounts && <Alert severity="info" sx={{ mb: 2 }}>La création de nouveaux comptes est désactivée ou indisponible dans cet environnement. Les comptes existants peuvent être gérés.</Alert>}
    {error && <Alert severity="error" sx={{ mb: 2 }} onClose={() => setError('')}>{error}</Alert>}
    {users.isPending || groups.isPending ? <CircularProgress /> : users.error || groups.error ? <Alert severity="error">{users.error?.message || groups.error?.message}</Alert> : <>
      <Paper sx={{ p: { xs: 2, md: 3 }, mb: 2 }}>
        <Box sx={{ display: 'flex', flexWrap: 'wrap', gap: 2, mb: 3 }}><TextField label="Rechercher un compte" value={search} onChange={e => setSearch(e.target.value)} disabled={busy} /><TextField select label={`Compte (${users.data.users.length})`} value={current?.username || ''} disabled={busy} sx={{ flex: 1, minWidth: 220 }} onChange={e => { const value = e.target.value; guard?.request(() => setSelected(value)) }}>{choices.map(p => <MenuItem key={p.username} value={p.username}>{p.username} · {p.role_tags.join(', ')}</MenuItem>)}</TextField></Box>
        {current ? <UserEditor key={current.username} token={token} user={current} users={users.data.users} groups={groups.data.groups} actor={session.username} busy={busy} run={run} /> : <Typography>Aucun compte enregistré.</Typography>}
      </Paper>
      <Accordion disableGutters><AccordionSummary expandIcon={<ExpandMoreRounded />}><Typography>Gérer les départements ({groups.data.groups.length})</Typography></AccordionSummary><AccordionDetails><GroupEditor token={token} groups={groups.data.groups} busy={busy} run={run} /></AccordionDetails></Accordion>
    </>}
    {creating && users.data && groups.data && <CreateAccount token={token} users={users.data.users} groups={groups.data.groups} enabled={capabilities.canCreateAccounts} busy={busy} run={run} error={error} onClose={() => setCreating(false)} />}
    <Snackbar open={!!message} autoHideDuration={5000} onClose={() => setMessage('')}><Alert severity="success" variant="filled" onClose={() => setMessage('')}>{message}</Alert></Snackbar>
  </Box>
}
