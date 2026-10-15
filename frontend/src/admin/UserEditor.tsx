import { useDraftGuard } from '../draftGuard'
import { useState } from 'react'
import { Accordion, AccordionDetails, AccordionSummary, Alert, Box, Button, Dialog, DialogActions, DialogContent, DialogTitle, MenuItem, TextField, Typography } from '@mui/material'
import ExpandMoreRounded from '@mui/icons-material/ExpandMoreRounded'
import type { Group } from '../team/api'
import { OrganizationFields } from './OrganizationFields'
import { adminApi } from './api'
import type { AdminUser, OrganizationInput, RunAction } from './api'
import { ContractEditor } from './ContractEditor'
import { DeleteAction } from '../components/DeleteAction'

export function UserEditor({ token, user, users, groups, actor, busy, run }: { token: string; user: AdminUser; users: AdminUser[]; groups: Group[]; actor: string; busy: boolean; run: RunAction }) {
  const [draft, setDraft] = useState<OrganizationInput>({ role_tags: user.role_tags, employee_type: user.employee_type, group_ids: user.group_ids, managed_group_ids: user.managed_group_ids, manager_id: user.manager_id })
  const [baseline, setBaseline] = useState(() => JSON.stringify(draft))
  const confirmChange = useDraftGuard(JSON.stringify(draft) !== baseline, busy, `Compte · ${user.username}`)
  const [remove, setRemove] = useState(false)
  const [confirmation, setConfirmation] = useState('')
  const [employee, setEmployee] = useState('')
  const reports = users.filter(p => p.manager_id === user.username)
  return <>
    <Typography variant="h6" sx={{ mb: 1, overflowWrap: 'anywhere' }}>Fiche de {user.username}</Typography>
    <Typography color="text.secondary" sx={{ mb: 2, overflowWrap: 'anywhere' }}>{user.email}</Typography>
    <Accordion disableGutters><AccordionSummary expandIcon={<ExpandMoreRounded />}><Typography>Fonction, départements et responsable</Typography></AccordionSummary><AccordionDetails>
      <Box component="form" onSubmit={e => { e.preventDefault(); void run(() => adminApi.organization(token, user.username, draft), 'Organisation enregistrée.').then(ok => { if (ok) setBaseline(JSON.stringify(draft)) }) }} sx={{ display: 'grid', gridTemplateColumns: { xs: '1fr', md: '1fr 1fr' }, gap: 2 }}>
        <OrganizationFields value={draft} onChange={setDraft} users={users} groups={groups} username={user.username} disabled={busy} />
        <Button variant="contained" type="submit" disabled={busy || !draft.role_tags.length}>Enregistrer l’organisation</Button>
      </Box>
    </AccordionDetails></Accordion>
    <Accordion disableGutters><AccordionSummary expandIcon={<ExpandMoreRounded />}><Typography>Contrats et historique</Typography></AccordionSummary><AccordionDetails><ContractEditor token={token} username={user.username} busy={busy} run={run} /></AccordionDetails></Accordion>
    <Accordion disableGutters><AccordionSummary expandIcon={<ExpandMoreRounded />}><Typography>Personnes sous responsabilité ({reports.length})</Typography></AccordionSummary><AccordionDetails>
      {reports.map(p => <Typography key={p.username} sx={{ overflowWrap: 'anywhere', mb: 1 }}>{p.username}</Typography>)}
      {!reports.length && <Typography color="text.secondary">Aucune personne rattachée.</Typography>}
      {user.role_tags.some(tag => ['Admin', 'Responsable'].includes(tag)) && <Box sx={{ display: 'flex', gap: 2, flexWrap: 'wrap', mt: 2 }}>
        <TextField select label="Affecter un employé existant" value={employee} disabled={busy} sx={{ minWidth: 230 }} onChange={e => setEmployee(e.target.value)}><MenuItem value="">Sélectionner</MenuItem>{users.filter(p => p.username !== user.username && p.role_tags.includes('Employe')).map(p => <MenuItem key={p.username} value={p.username}>{p.username} · {p.manager_id || 'Sans responsable'}</MenuItem>)}</TextField>
        <Button disabled={busy || !employee} onClick={() => void run(() => adminApi.manager(token, employee, user.username), 'Employé affecté au responsable.')}>Affecter</Button>
      </Box>}
    </AccordionDetails></Accordion>
    {user.username !== actor && <Box sx={{ mt: 2 }}><DeleteAction label="Supprimer le compte" disabled={busy} onClick={() => confirmChange(() => setRemove(true))} /></Box>}
    <Dialog open={remove} onClose={() => { if (!busy) setRemove(false) }}><DialogTitle>Supprimer le compte et ses données associées ?</DialogTitle><DialogContent>
      <Alert severity="warning" sx={{ mb: 2 }}>Cette opération supprime le compte, ses contrats, les fiches associées et les documents liés, y compris ceux déposés pour d’autres personnes. Elle est définitive.</Alert>
      <TextField fullWidth label={`Saisissez ${user.username} pour confirmer`} value={confirmation} onChange={e => setConfirmation(e.target.value)} disabled={busy} />
    </DialogContent><DialogActions><Button disabled={busy} onClick={() => setRemove(false)}>Annuler</Button><Button disabled={busy || confirmation !== user.username} onClick={() => void run(() => adminApi.remove(token, user.username), 'Compte supprimé.')}>Supprimer définitivement</Button></DialogActions></Dialog>
  </>
}
