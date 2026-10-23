import { useState } from 'react'
import { Alert, Box, Button, Dialog, DialogActions, DialogContent, DialogTitle, TextField } from '@mui/material'
import { useDraftGuard } from '../draftGuard'
import { request } from '../api'
import { OrganizationFields } from './OrganizationFields'
import type { AdminUser, OrganizationInput, RunAction } from './api'
import type { Group } from '../team/api'

const initial = (): OrganizationInput => ({ role_tags: ['Employe'], employee_type: 'salarie', manager_id: null, group_ids: [], managed_group_ids: [] })
export function CreateAccount({ token, users, groups, allowSuperAdmin, enabled, busy, run, error, onClose }: { token: string; users: AdminUser[]; groups: Group[]; allowSuperAdmin: boolean; enabled: boolean; busy: boolean; run: RunAction; error: string; onClose: () => void }) {
  const [firstName, setFirstName] = useState('')
  const [lastName, setLastName] = useState('')
  const [nameError, setNameError] = useState('')
  const [email, setEmail] = useState('')
  const [password, setPassword] = useState('')
  const [confirmation, setConfirmation] = useState('')
  const [organization, setOrganization] = useState(initial)
  const close = useDraftGuard(!!firstName || !!lastName || !!email || !!password || !!confirmation || JSON.stringify(organization) !== JSON.stringify(initial()), busy, 'Nouveau compte')
  const disabled = busy || !enabled
  async function create() {
    if (disabled || password !== confirmation || password.length < 8) return
    const first_name = firstName.trim().replace(/\s+/g, ' ')
    const last_name = lastName.trim().replace(/\s+/g, ' ')
    if (!first_name || !last_name || `${first_name} ${last_name}`.length > 200) { setNameError('Les prénoms et le nom sont obligatoires (200 caractères maximum au total).'); return }
    setNameError('')
    const ok = await run(() => request('/create-user', token, { method: 'POST', body: JSON.stringify({ first_name, last_name, new_mail: email.trim(), new_password: password, ...organization }) }), 'Compte créé.')
    if (ok) onClose()
  }
  return <Dialog open fullWidth maxWidth="md" aria-labelledby="create-account-title" onClose={() => { if (!busy) close(onClose) }}>
    <Box component="form" onSubmit={e => { e.preventDefault(); void create() }}>
    <DialogTitle id="create-account-title">Créer un compte</DialogTitle>
    <DialogContent>
    {error && <Alert severity="error" sx={{ mb: 2 }}>{error}</Alert>}
    {nameError && <Alert severity="error" sx={{ mb: 2 }}>{nameError}</Alert>}
    {!enabled && <Alert severity="info" sx={{ mb: 2 }}>Le formulaire est disponible, mais la création de comptes est désactivée ou indisponible dans cet environnement.</Alert>}
    <Box sx={{ display: 'grid', gridTemplateColumns: { xs: '1fr', md: '1fr 1fr' }, gap: 2, pt: 1 }}>
      <TextField label="Prénoms" autoComplete="given-name" required value={firstName} disabled={disabled} onChange={e => setFirstName(e.target.value)} slotProps={{ htmlInput: { maxLength: 100 } }} />
      <TextField label="Nom" autoComplete="family-name" required value={lastName} disabled={disabled} onChange={e => setLastName(e.target.value)} slotProps={{ htmlInput: { maxLength: 100 } }} />
      <TextField label="E-mail du nouveau compte" type="email" required value={email} disabled={disabled} onChange={e => setEmail(e.target.value)} sx={{ gridColumn: '1 / -1' }} />
      <TextField label="Mot de passe initial" type="password" autoComplete="new-password" required value={password} disabled={disabled} onChange={e => setPassword(e.target.value)} slotProps={{ htmlInput: { minLength: 8 } }} />
      <TextField label="Confirmer le mot de passe initial" type="password" autoComplete="new-password" required value={confirmation} disabled={disabled} error={!!confirmation && password !== confirmation} onChange={e => setConfirmation(e.target.value)} />
      <OrganizationFields value={organization} onChange={setOrganization} users={users} groups={groups} disabled={disabled} allowSuperAdmin={allowSuperAdmin} />
    </Box>
    </DialogContent>
    <DialogActions sx={{ px: 3, pb: 3 }}>
      <Button disabled={busy} onClick={() => close(onClose)}>Annuler</Button>
      <Button type="submit" variant="contained" disabled={disabled || !firstName.trim() || !lastName.trim() || password !== confirmation || !organization.role_tags.length}>{busy ? 'Création…' : 'Créer le compte'}</Button>
    </DialogActions>
    </Box>
  </Dialog>
}
