import { personDisplayName } from '../personDisplayName'
import { isAdmin } from '../rolePolicy'
import { AccountList } from './AccountList'
import { accountStatus, type AccountStatusFilter } from './accountStatus'
import { selectedAccount } from './accountSelection'
import { AdvancedMaintenance } from './AdvancedMaintenance'
import { useDraftGuard, useWorkspaceGuard } from '../draftGuard'
import { useId, useRef, useState } from 'react'
import { useWorkNavigation } from '../workspaceNavigation'
import { Alert, Box, Button, CircularProgress, Dialog, DialogActions, DialogContent, DialogTitle, IconButton, Snackbar, Tab, Tabs, TextField } from '@mui/material'
import GroupsOutlined from '@mui/icons-material/GroupsOutlined'
import CloseRounded from '@mui/icons-material/CloseRounded'
import SaveOutlined from '@mui/icons-material/SaveOutlined'
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
  const navigation = useWorkNavigation()
  const client = useQueryClient()
  const users = useQuery({ queryKey: ['admin-users', session.username], queryFn: ({ signal }) => adminApi.users(token, signal), retry: false })
  const groups = useQuery({ queryKey: ['admin-groups', session.username], queryFn: ({ signal }) => adminApi.groups(token, signal), retry: false })
  const [selected, setSelected] = useState('')
  const [status, setStatus] = useState<AccountStatusFilter>('active')
  const statusId = useId()
  const organizationFormId = `${statusId}-organization`
  const [search, setSearch] = useState('')
  const [managingGroups, setManagingGroups] = useState(false)
  const [appliedWork, setAppliedWork] = useState<number | null>(null)
  if (navigation?.request?.page === 'admin' && appliedWork !== navigation.request.sequence) {
    setAppliedWork(navigation.request.sequence)
    setSelected(navigation.request.username || ''); setSearch(''); setStatus('all')
  }
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
  const accounts = (users.data?.users || []).filter(person => status === 'all' || accountStatus(person) === status)
  const current = selected ? selectedAccount(users.data?.users || [], selected) : undefined
  const choices = accounts.filter(p => `${personDisplayName(p)} ${p.email || ''}`.toLocaleLowerCase('fr').includes(search.trim().toLocaleLowerCase('fr')))
  function closeAccount() { guard?.request(() => { setSelected(''); setError('') }) }
  return <Box className="workspace">
    <WorkspaceToolbar>
      <Box className="workspace-action-bar" sx={{ width: '100%' }}>
        <Tabs value={status} aria-label="État des comptes" variant="scrollable" scrollButtons="auto" sx={{ minHeight: 40, maxWidth: '100%', '& .MuiTab-root': { minHeight: 40, minWidth: 0, px: 1.5, fontSize: 13 } }}
          onChange={(_, next: AccountStatusFilter) => guard?.request(() => { setStatus(next); setSelected(''); setSearch('') })}>
          {([['active', 'Normal'], ['inactive', 'Désactivés'], ['archived', 'Archivés'], ['deleted', 'Connexion supprimée'], ['all', 'Tous']] as const).map(([value, label]) =>
            <Tab key={value} value={value} label={label} disabled={busy} id={`${statusId}-${value}`} aria-controls={`${statusId}-panel`} />)}
        </Tabs>
        <TextField size="small" placeholder="Rechercher un compte" value={search} onChange={event => setSearch(event.target.value)} disabled={busy || !users.data || !groups.data} slotProps={{ htmlInput: { 'aria-label': 'Rechercher un compte' } }} sx={{ width: { xs: '100%', sm: 220 }, maxWidth: '100%' }} />
        <Button className="workspace-primary-action" variant="contained" startIcon={<GroupsOutlined />} disabled={busy || !groups.data || !!groups.error} onClick={() => setManagingGroups(true)} sx={{ ml: 'auto', whiteSpace: 'nowrap' }}>Gérer les départements</Button>
      </Box>
    </WorkspaceToolbar>
    {error && <Alert severity="error" sx={{ mb: 2 }} onClose={() => setError('')}>{error}</Alert>}
    {users.isPending || groups.isPending ? <CircularProgress /> : users.error || groups.error ? <Alert severity="error">{users.error?.message || groups.error?.message}</Alert> : <>
      <Box role="tabpanel" id={`${statusId}-panel`} aria-labelledby={`${statusId}-${status}`}>
        <AccountList accounts={choices} busy={busy} onOpen={username => { setSelected(username); setError('') }} />
      </Box>
      {!!selected && <Dialog open fullWidth maxWidth="lg" onClose={busy ? undefined : closeAccount}
        aria-labelledby={`${statusId}-detail-title`} slotProps={{ paper: { sx: { maxHeight: 'min(820px, 90dvh)', m: { xs: 2, sm: 4 }, width: { xs: 'calc(100% - 32px)', sm: 'calc(100% - 64px)' } } } }}>
        <DialogTitle id={`${statusId}-detail-title`} sx={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', py: 1.5 }}>
          Détails du compte<IconButton aria-label="Fermer le compte" disabled={busy} onClick={closeAccount}><CloseRounded /></IconButton>
        </DialogTitle>
        <DialogContent dividers>
          {error && <Alert severity="error" sx={{ mb: 2 }} onClose={() => setError('')}>{error}</Alert>}
          {current ? <UserEditor key={current.username} completion={navigation?.request?.page === 'admin' && navigation.request.username === current.username ? navigation.request : undefined} formId={organizationFormId} token={token} user={current} users={users.data.users} groups={groups.data.groups} actor={session.username} actorRoles={session.role_tags} busy={busy} run={run} requestError={error} /> : <Alert severity="warning">Le compte demandé est introuvable ou inaccessible.</Alert>}
        </DialogContent>
        <DialogActions sx={{ px: 3, py: 1.5, flexShrink: 0 }}>
          <Button disabled={busy} onClick={closeAccount}>Fermer</Button>
          <Button variant="contained" startIcon={<SaveOutlined />} type="submit" form={organizationFormId} disabled={busy || !current}>{busy ? 'Enregistrement…' : 'Enregistrer'}</Button>
        </DialogActions>
      </Dialog>}
    </>}
    <AdvancedMaintenance token={token} session={session} />
    {managingGroups && groups.data && <GroupEditor token={token} groups={groups.data.groups} busy={busy} run={run} onClose={() => setManagingGroups(false)} />}
    <Snackbar open={!!message} autoHideDuration={5000} onClose={() => setMessage('')}><Alert severity="success" variant="filled" onClose={() => setMessage('')}>{message}</Alert></Snackbar>
  </Box>
}
