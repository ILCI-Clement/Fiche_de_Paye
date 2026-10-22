import { personDisplayName } from '../personDisplayName'
import { isAdmin, isSuperAdmin } from '../rolePolicy'
import { AccountOption, AccountOptionHeader } from './AccountOption'
import { accountMenuProps } from './accountOptionLayout'
import { AdvancedMaintenance } from './AdvancedMaintenance'
import { useCapabilities } from '../capabilities'
import { CreateAccount } from './CreateAccount'
import { useDraftGuard, useWorkspaceGuard } from '../draftGuard'
import { useRef, useState } from 'react'
import { Accordion, AccordionDetails, AccordionSummary, Alert, Box, Button, CircularProgress, IconButton, ListSubheader, MenuItem, Paper, Snackbar, TextField, Tooltip, Typography } from '@mui/material'
import ExpandMoreRounded from '@mui/icons-material/ExpandMoreRounded'
import PersonAddOutlined from '@mui/icons-material/PersonAddOutlined'
import RefreshRounded from '@mui/icons-material/RefreshRounded'
import { useQuery, useQueryClient } from '@tanstack/react-query'
import type { Session } from '../types'
import { adminApi } from './api'
import { GroupEditor } from './GroupEditor'
import { UserEditor } from './UserEditor'
import { WorkspaceToolbar } from '../components/WorkspaceToolbar'

export function AdminPage({ token, session }: { token: string; session: Session }) {
  if (!isAdmin(session)) return <Box className="workspace"><Alert severity="error">Cette page est réservée aux administrateurs.</Alert></Box>
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
  const choices = (users.data?.users || []).filter(p => p.username === current?.username || `${personDisplayName(p)} ${p.email}`.toLocaleLowerCase('fr').includes(search.toLocaleLowerCase('fr')))
  return <Box className="workspace">
    <WorkspaceToolbar>
      <Box className="workspace-action-bar" sx={{ width: '100%' }}>
        <TextField size="small" placeholder="Rechercher un compte" value={search} onChange={event => setSearch(event.target.value)} disabled={busy || !users.data || !groups.data} slotProps={{ htmlInput: { 'aria-label': 'Rechercher un compte' } }} sx={{ width: { xs: '100%', sm: 220 }, maxWidth: '100%' }} />
        <TextField select size="small" label={`Compte (${users.data?.users.length || 0})`} value={current?.username || ''} disabled={busy || !users.data || !groups.data} slotProps={{ select: { renderValue: () => current ? <AccountOption person={current} compact /> : '', MenuProps: accountMenuProps } }} sx={{ minWidth: 'min(100%, 220px)', maxWidth: 360, flex: { xs: '1 1 100%', sm: '1 1 260px' } }} onChange={event => { const value = event.target.value; guard?.request(() => setSelected(value)) }}>
          <ListSubheader><AccountOptionHeader /></ListSubheader>
          {choices.map(person => <MenuItem key={person.username} value={person.username}><AccountOption person={person} /></MenuItem>)}
        </TextField>
        <Button variant="outlined" startIcon={<PersonAddOutlined />} disabled={busy || !users.data || !groups.data || !!users.error || !!groups.error} onClick={() => { setError(''); setCreating(true) }}>Créer un compte</Button>
        <Tooltip title="Actualiser"><span><IconButton aria-label="Actualiser" disabled={busy} onClick={() => { guard?.request(() => { void users.refetch(); void groups.refetch() }) }}><RefreshRounded /></IconButton></span></Tooltip>
      </Box>
    </WorkspaceToolbar>
    {!capabilities.canCreateAccounts && <Alert severity="info" sx={{ mb: 2 }}>La création de nouveaux comptes est désactivée ou indisponible dans cet environnement. Les comptes existants peuvent être gérés.</Alert>}
    {error && <Alert severity="error" sx={{ mb: 2 }} onClose={() => setError('')}>{error}</Alert>}
    {users.isPending || groups.isPending ? <CircularProgress /> : users.error || groups.error ? <Alert severity="error">{users.error?.message || groups.error?.message}</Alert> : <>
      <Paper sx={{ p: { xs: 2, md: 2.5 }, mb: 2 }}>
        {current ? <UserEditor key={current.username} token={token} user={current} users={users.data.users} groups={groups.data.groups} actor={session.username} actorRoles={session.role_tags} busy={busy} run={run} requestError={error} /> : <Typography>Aucun compte enregistré.</Typography>}
      </Paper>
      <Accordion disableGutters><AccordionSummary expandIcon={<ExpandMoreRounded />}><Typography>Gérer les départements ({groups.data.groups.length})</Typography></AccordionSummary><AccordionDetails><GroupEditor token={token} groups={groups.data.groups} busy={busy} run={run} /></AccordionDetails></Accordion>
    </>}
    <AdvancedMaintenance token={token} session={session} />
    {creating && users.data && groups.data && <CreateAccount token={token} users={users.data.users} groups={groups.data.groups} allowSuperAdmin={isSuperAdmin(session)} enabled={capabilities.canCreateAccounts} busy={busy} run={run} error={error} onClose={() => setCreating(false)} />}
    <Snackbar open={!!message} autoHideDuration={5000} onClose={() => setMessage('')}><Alert severity="success" variant="filled" onClose={() => setMessage('')}>{message}</Alert></Snackbar>
  </Box>
}
