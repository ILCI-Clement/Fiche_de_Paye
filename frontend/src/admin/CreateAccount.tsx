import { useState } from 'react'
import { Alert, Box, Button, TextField } from '@mui/material'
import { useDraftGuard } from '../draftGuard'
import { request } from '../api'
import { OrganizationFields } from './OrganizationFields'
import type { AdminUser, OrganizationInput, RunAction } from './api'
import type { Group } from '../team/api'

const initial = (): OrganizationInput => ({ role_tags: ['Employe'], employee_type: 'salarie', manager_id: null, group_ids: [], managed_group_ids: [] })
export function CreateAccount({ token, users, groups, enabled, busy, run }: { token: string; users: AdminUser[]; groups: Group[]; enabled: boolean; busy: boolean; run: RunAction }) {
  const [name, setName] = useState('')
  const [email, setEmail] = useState('')
  const [password, setPassword] = useState('')
  const [confirmation, setConfirmation] = useState('')
  const [organization, setOrganization] = useState(initial)
  useDraftGuard(!!name || !!email || !!password || !!confirmation || JSON.stringify(organization) !== JSON.stringify(initial()), busy, 'Nouveau compte')
  const disabled = busy || !enabled
  async function create() {
    if (disabled || password !== confirmation || password.length < 8) return
    const ok = await run(() => request('/create-user', token, { method: 'POST', body: JSON.stringify({ new_username: name.trim(), new_mail: email.trim(), new_password: password, ...organization }) }), 'Compte créé.')
    if (ok) { setName(''); setEmail(''); setPassword(''); setConfirmation(''); setOrganization(initial()) }
  }
  return <>
    {!enabled && <Alert severity="info" sx={{ mb: 2 }}>Le formulaire est disponible, mais la création de comptes est désactivée ou indisponible dans cet environnement.</Alert>}
    <Box component="form" onSubmit={e => { e.preventDefault(); void create() }} sx={{ display: 'grid', gridTemplateColumns: { xs: '1fr', md: '1fr 1fr' }, gap: 2 }}>
      <TextField label="Nom du nouveau compte" autoComplete="off" required value={name} disabled={disabled} onChange={e => setName(e.target.value)} />
      <TextField label="E-mail du nouveau compte" type="email" required value={email} disabled={disabled} onChange={e => setEmail(e.target.value)} />
      <TextField label="Mot de passe initial" type="password" autoComplete="new-password" required value={password} disabled={disabled} onChange={e => setPassword(e.target.value)} slotProps={{ htmlInput: { minLength: 8 } }} />
      <TextField label="Confirmer le mot de passe initial" type="password" autoComplete="new-password" required value={confirmation} disabled={disabled} error={!!confirmation && password !== confirmation} onChange={e => setConfirmation(e.target.value)} />
      <OrganizationFields value={organization} onChange={setOrganization} users={users} groups={groups} disabled={disabled} />
      <Button type="submit" variant="contained" disabled={disabled || password !== confirmation || !organization.role_tags.length}>Créer le compte</Button>
    </Box>
  </>
}
