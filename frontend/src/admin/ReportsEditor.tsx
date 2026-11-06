import { useState } from 'react'
import { Box, Button, Checkbox, Chip, Dialog, DialogActions, DialogContent, DialogTitle, Alert, TextField, Typography } from '@mui/material'
import { DepartmentBadges } from '../components/DepartmentBadge'
import type { Group } from '../team/api'
import { useDraftGuard } from '../draftGuard'
import { canManageTeam, isSuperAdmin } from '../rolePolicy'
import { personDisplayName } from '../personDisplayName'
import { accountStatus } from './accountStatus'
import { adminApi, type AdminUser, type RunAction } from './api'
import { reportChanges } from './reportChanges'

export function ReportsEditor({ token, manager, users, groups, disabled, busy, run }: {
  token: string; manager: AdminUser; users: AdminUser[]; groups: Group[]; disabled: boolean; busy: boolean; run: RunAction;
}) {
  const [editing, setEditing] = useState(false)
  const reports = users.filter(person => person.manager_id === manager.username && accountStatus(person) === 'active')
  return <>
    <Box sx={{ display: 'flex', justifyContent: 'flex-end', mb: 1 }}>
      {!disabled && ((canManageTeam(manager) && accountStatus(manager) === 'active') || reports.length > 0) && <Button size="small" disabled={busy}
        onClick={() => setEditing(true)}>Modifier la liste</Button>}
    </Box>
    <Box sx={{ display: 'grid', gap: 1 }}>
      {reports.map(person => {
        const departments = groups.filter(group => person.group_ids.includes(group.id)).map(group => group.name)
        const name = personDisplayName(person)
        return <Box key={person.username} sx={{ display: 'grid', gridTemplateColumns: { xs: 'minmax(0, 1fr) minmax(0, 1fr)', sm: '240px minmax(0, 1fr)' }, alignItems: 'center', gap: 1.5, minWidth: 0 }}>
          <Chip label={name} title={name} variant="outlined" sx={{ borderRadius: 1.5, fontWeight: 500, width: '100%', minWidth: 0, height: 32, justifyContent: 'flex-start', '& .MuiChip-label': { overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' } }} />
          {departments.length ? <Box sx={{ minWidth: 0, maxWidth: '100%', '& .department-badges': { display: 'flex', flexDirection: 'row', flexWrap: 'wrap', gap: .75 },
            '& .department-badge': { width: 200, maxWidth: '100%', minWidth: 0 } }}><DepartmentBadges names={departments} /></Box>
            : <Typography variant="caption" color="text.secondary">Sans département</Typography>}
        </Box>
      })}
    </Box>
    {!reports.length && <Typography color="text.secondary" variant="body2">Aucune personne rattachée.</Typography>}
    {editing && <ReportsDialog token={token} manager={manager} users={users} busy={busy} run={run} onClose={() => setEditing(false)} />}
  </>
}

function ReportsDialog({ token, manager, users, busy, run, onClose }: {
  token: string; manager: AdminUser; users: AdminUser[]; busy: boolean; run: RunAction; onClose: () => void;
}) {
  // Freeze identities and prior assignments so an external edit cannot silently alter the submitted baseline.
  const [accounts] = useState(() => users.filter(person => accountStatus(person) === 'active'
    && person.username !== manager.username && (!isSuperAdmin(person) || person.manager_id === manager.username)).map(person => ({ ...person })))
  const [selected, setSelected] = useState(() => accounts.filter(person => person.manager_id === manager.username).map(person => person.username))
  const [search, setSearch] = useState('')
  const [confirm, setConfirm] = useState(false)
  const [error, setError] = useState('')
  const changes = reportChanges(accounts, manager.username, selected)
  const assignable = canManageTeam(manager) && accountStatus(manager) === 'active'
  const close = useDraftGuard(changes.length > 0, busy, 'Personnes sous responsabilité')
  const shown = accounts.filter(person => `${personDisplayName(person)} ${person.email}`.toLocaleLowerCase('fr').includes(search.trim().toLocaleLowerCase('fr')))
  const managerName = (id: string | null) => id ? personDisplayName(users.find(person => person.username === id)
    || { display_name: accounts.find(person => person.manager_id === id)?.manager_display_name || 'Responsable actuel' }) : 'Sans responsable'
  async function save() {
    if (!changes.length || busy) return
    setError('')
    const ok = await run(async () => {
      try { return await adminApi.reports(token, manager.username, changes) }
      catch (failure) { setError(failure instanceof Error ? failure.message : 'Enregistrement impossible.'); throw failure }
    }, 'Liste enregistrée.')
    if (ok) { setConfirm(false); onClose() }
  }
  return <>
    <Dialog open={!confirm} fullWidth maxWidth="md" onClose={busy ? undefined : () => close(onClose)}>
      <DialogTitle>Équipe · {personDisplayName(manager)}</DialogTitle>
      <DialogContent dividers sx={{ maxHeight: '65dvh' }}>
        <TextField autoFocus fullWidth size="small" placeholder="Nom ou e-mail" value={search} disabled={busy}
          slotProps={{ htmlInput: { 'aria-label': 'Rechercher une personne' } }} onChange={event => setSearch(event.target.value)} sx={{ mb: 1.5 }} />
        <Typography variant="caption" color="text.secondary">{selected.length} personne(s) sélectionnée(s)</Typography>
        {shown.map(person => <Box key={person.username} component="label" sx={{ display: 'flex', gap: 1, alignItems: 'center', py: .5, borderBottom: 1, borderColor: 'divider', cursor: 'pointer' }}>
          <Checkbox checked={selected.includes(person.username)} disabled={busy || (!assignable && person.manager_id !== manager.username)}
            slotProps={{ input: { 'aria-label': personDisplayName(person) } }}
            onChange={(_, checked) => setSelected(current => checked ? [...current, person.username] : current.filter(name => name !== person.username))} />
          <Box sx={{ flex: 1, minWidth: 0 }}><Typography variant="body2">{personDisplayName(person)}</Typography>
            <Typography variant="caption" color="text.secondary" sx={{ overflowWrap: 'anywhere' }}>{person.email}</Typography></Box>
          <Typography variant="caption" color="text.secondary" sx={{ maxWidth: '35%', overflowWrap: 'anywhere' }}>{managerName(person.manager_id)}</Typography>
        </Box>)}
        {!shown.length && <Typography color="text.secondary" sx={{ py: 2 }}>Aucune personne trouvée.</Typography>}
        {changes.length > 100 && <Alert severity="warning" sx={{ mt: 1 }}>100 modifications maximum par enregistrement.</Alert>}
      </DialogContent>
      <DialogActions><Button disabled={busy} onClick={() => close(onClose)}>Annuler</Button>
        <Button variant="contained" disabled={busy || !changes.length || changes.length > 100} onClick={() => { setError(''); setConfirm(true) }}>Vérifier ({changes.length})</Button></DialogActions>
    </Dialog>
    <Dialog open={confirm} fullWidth maxWidth="md" onClose={busy ? undefined : () => setConfirm(false)}>
      <DialogTitle>Confirmer les affectations</DialogTitle>
      <DialogContent dividers sx={{ maxHeight: '65dvh' }}>
        {error && <Alert severity="error" sx={{ mb: 1 }}>{error}</Alert>}
        {changes.map(change => <Box key={change.username} sx={{ py: 1, borderBottom: 1, borderColor: 'divider' }}>
          <Typography variant="body2" sx={{ fontWeight: 500 }}>{personDisplayName(accounts.find(person => person.username === change.username))}</Typography>
          <Typography variant="body2" color="text.secondary">{managerName(change.expected_manager_id)} → {change.assigned ? personDisplayName(manager) : 'Sans responsable'}</Typography>
        </Box>)}
      </DialogContent>
      <DialogActions><Button disabled={busy} onClick={() => setConfirm(false)}>Revenir</Button>
        <Button variant="contained" disabled={busy} onClick={() => void save()}>{busy ? 'Enregistrement…' : 'Confirmer'}</Button></DialogActions>
    </Dialog>
  </>
}
