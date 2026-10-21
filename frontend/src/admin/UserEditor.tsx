import { useDraftGuard } from '../draftGuard'
import { useState } from 'react'
import { Accordion, AccordionDetails, AccordionSummary, Alert, Box, Button, Dialog, DialogActions, DialogContent, DialogTitle, MenuItem, TextField, Typography } from '@mui/material'
import ExpandMoreRounded from '@mui/icons-material/ExpandMoreRounded'
import type { Group } from '../team/api'
import { OrganizationFields } from './OrganizationFields'
import { adminApi } from './api'
import type { AdminUser, OrganizationInput, RunAction } from './api'
import { ContractEditor } from './ContractEditor'
import { AccountAccess } from './AccountAccess'

export function UserEditor({ token, user, users, groups, actor, busy, run, requestError = '' }: { token: string; user: AdminUser; users: AdminUser[]; groups: Group[]; actor: string; busy: boolean; run: RunAction; requestError?: string }) {
  const [draft, setDraft] = useState<OrganizationInput>({ role_tags: user.role_tags, employee_type: user.employee_type, group_ids: user.group_ids, managed_group_ids: user.managed_group_ids, manager_id: user.manager_id })
  const [baseline, setBaseline] = useState(() => JSON.stringify(draft))
  const confirmChange = useDraftGuard(JSON.stringify(draft) !== baseline, busy, `Compte · ${user.username}`)
  const [employee, setEmployee] = useState('')
  const [roleConfirm, setRoleConfirm] = useState(false)
  async function saveOrganization() {
    const ok = await run(() => adminApi.organization(token, user.username, draft), 'Organisation enregistrée.')
    if (ok) { setBaseline(JSON.stringify(draft)); setRoleConfirm(false) }
  }
  const reports = users.filter(p => p.manager_id === user.username)
  return <>
    <Typography variant="h6" sx={{ mb: 1, overflowWrap: 'anywhere' }}>Fiche de {user.username}</Typography>
    <Typography color="text.secondary" sx={{ mb: 2, overflowWrap: 'anywhere' }}>{user.email}</Typography>
    <Accordion disableGutters><AccordionSummary expandIcon={<ExpandMoreRounded />}><Typography>Fonction, départements et responsable</Typography></AccordionSummary><AccordionDetails>
      <Box component="form" onSubmit={e => { e.preventDefault(); if ([...draft.role_tags].sort().join() !== [...user.role_tags].sort().join()) setRoleConfirm(true); else void saveOrganization() }} sx={{ display: 'grid', gridTemplateColumns: { xs: '1fr', md: '1fr 1fr' }, gap: 2 }}>
        <OrganizationFields value={draft} onChange={setDraft} users={users} groups={groups} username={user.username} disabled={busy} preserveAdmin={user.username === actor && user.role_tags.includes('Admin')} />
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
    <AccountAccess token={token} user={user} actor={actor} busy={busy} run={run} confirmChange={confirmChange} requestError={requestError} />
    <Dialog open={roleConfirm} onClose={() => { if (!busy) setRoleConfirm(false) }}><DialogTitle>Confirmer le changement de droits</DialogTitle><DialogContent>
      <Alert severity="warning" sx={{ mb: 2 }}>Les accès de {user.username} seront modifiés.</Alert>
      <Typography>Actuellement : {user.role_tags.join(' · ')}</Typography>
      <Typography>Après modification : {draft.role_tags.join(' · ')}</Typography>
    </DialogContent><DialogActions><Button disabled={busy} onClick={() => setRoleConfirm(false)}>Annuler</Button><Button variant="contained" disabled={busy} onClick={() => void saveOrganization()}>Confirmer les nouveaux droits</Button></DialogActions></Dialog>
  </>
}
