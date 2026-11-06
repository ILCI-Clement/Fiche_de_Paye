import { useRef, useState } from 'react'
import { missingPersonnelField } from '../completionTarget'
import { useCompletionFocus } from '../components/useCompletionFocus'
import { Alert, Box, Button, Checkbox, Dialog, DialogActions, DialogContent, DialogTitle, Divider, FormControlLabel, IconButton, InputAdornment, MenuItem, TextField, Tooltip, Typography } from '@mui/material'
import VisibilityOutlined from '@mui/icons-material/VisibilityOutlined'
import VisibilityOffOutlined from '@mui/icons-material/VisibilityOffOutlined'
import { SensitiveNumber } from './SensitiveNumber'
import { useDraftGuard } from '../draftGuard'
import { FrenchDateField } from '../components/FrenchDateField'
import { personnelApi, personnelTypes, blankPeriod } from './api'
import type { Employee, PersonnelInput, PersonnelType, PersonnelWriteInput } from './api'
import { AccountLink } from './AccountLink'
import { EmploymentFields } from './EmploymentFields'
import { useRegistryAction } from './useRegistryAction'
import { WriteError } from './WriteError'
import { useSensitiveNumberCapability } from './useSensitiveNumberCapability'
import { missingRegistrationFields, personnelFieldLabels, validateEmployment } from './personnelValidation'
import { RegistrationContract } from './RegistrationContract'
import { useQuery, useQueryClient } from '@tanstack/react-query'
import { useRegistryViewer } from './context'
import { adminApi, type OrganizationInput } from '../admin/api'
import { organizationValidation } from '../admin/organizationValidation'
import { RegistrationAccountFields } from './RegistrationAccountFields'
import { useCapabilities } from '../capabilities'

const initialOrganization = (): OrganizationInput => ({ role_tags: ['Employe'], employee_type: 'salarie', manager_id: null, group_ids: [], managed_group_ids: [] })

function initialInput(employee?: Employee): PersonnelInput {
  return { full_name: employee?.full_name || '', last_name: employee?.last_name || '', first_name: employee?.first_name || '', nationality: employee?.nationality || '', sex: employee?.sex || '', address: employee?.address || '', birth_date: employee?.birth_date || null, job_title: employee?.job_title || '', qualification: employee?.qualification || '', personnel_type: employee?.personnel_type || '', linked_username: employee?.linked_username || null }
}

export function PersonnelEditor({ token, employee, onClose, onSaved, allowSuperAdmin = false, completionCode }: { completionCode?: string; token: string; employee?: Employee; onClose: () => void; onSaved: (message: string) => void; allowSuperAdmin?: boolean }) {
  const completionRoot = useRef<HTMLDivElement>(null)
  useCompletionFocus(completionRoot, employee && completionCode ? missingPersonnelField(employee as unknown as Record<string, unknown>, completionCode) : undefined)
  const client = useQueryClient()
  const [version] = useState(employee?.version)
  const [value, setValue] = useState(() => initialInput(employee))
  const [linkedDisplayName, setLinkedDisplayName] = useState(employee?.linked_display_name || '')
  const [initial] = useState(() => JSON.stringify(initialInput(employee)))
  const [period, setPeriod] = useState(blankPeriod)
  const [validation, setValidation] = useState('')
  const [socialSecurityNumber, setSocialSecurityNumber] = useState('')
  const [numberVisible, setNumberVisible] = useState(false)
  const [clearSocialSecurityNumber, setClearSocialSecurityNumber] = useState(false)
  const [email, setEmail] = useState(employee?.email || '')
  const [checkedEmail, setCheckedEmail] = useState('')
  const [createAccount, setCreateAccount] = useState(false)
  const [external, setExternal] = useState(false)
  const wantsAccount = createAccount || external
  const checksLoginEmail = wantsAccount || !!employee?.linked_username
  const viewer = useRegistryViewer()
  const emailLookup = useQuery({
    queryKey: ['personnel-email', viewer, employee?.employee_id, checkedEmail],
    queryFn: ({ signal }) => personnelApi.emailAvailability(token, checkedEmail, employee?.employee_id, signal),
    enabled: !!checkedEmail && checkedEmail === email.trim() && checksLoginEmail,
    retry: false, gcTime: 0,
  })
  const emailTaken = checksLoginEmail && checkedEmail === email.trim() && emailLookup.data?.available === false
  const [organization, setOrganization] = useState(initialOrganization)
  const [createdAccount, setCreatedAccount] = useState<{ welcome_email_sent: boolean } | null>(null)
  const capabilities = useCapabilities()
  const capability = useSensitiveNumberCapability(token)
  const nirEnabled = capability.data?.nir_enabled === true && !capability.error
  const action = useRegistryAction()
  const strictRegistration = !employee || employee.registration_schema_version >= 2
  const dirty = JSON.stringify(value) !== initial || JSON.stringify(period) !== JSON.stringify(blankPeriod()) || !!socialSecurityNumber || clearSocialSecurityNumber
  const close = useDraftGuard(!createdAccount && (dirty || email !== (employee?.email || '') || wantsAccount), action.busy, employee ? 'Fiche du personnel' : 'Nouvelle inscription au registre')
  const field = (key: keyof PersonnelInput, next: string | null) => setValue(current => {
    const updated = { ...current, [key]: next }
    if (updated.first_name.trim() && updated.last_name.trim()) updated.full_name = `${updated.first_name.trim()} ${updated.last_name.trim()}`
    return updated
  })
  async function save() {
    if (action.busy) return
    if (createdAccount) return
    if (emailTaken) { setValidation('Cette adresse e-mail est déjà utilisée par un compte.'); return }
    if (wantsAccount) {
      const error = organizationValidation(organization)
      if (!capabilities.canCreateAccounts || !email.trim() || error) { setValidation(error || 'Renseignez un e-mail valide et vérifiez la disponibilité de la création de comptes.'); return }
    }
    if (!value.full_name.trim()) { setValidation('Renseignez le nom et les prénoms.'); return }
    if (external && (!value.first_name.trim() || !value.last_name.trim())) { setValidation('Renseignez les prénoms et le nom.'); return }
    if (!external && strictRegistration) {
      const missing = missingRegistrationFields(value, employee ? { ...blankPeriod(), entry_date: employee.latest_period?.entry_date || null } : period)
      if (missing.length) { setValidation(`Champs obligatoires : ${missing.join(', ')}.`); return }
    }
    if (!external && !employee) {
      const periodError = validateEmployment(period)
      if (periodError) { setValidation(periodError); return }
      if (value.birth_date && period.entry_date! < value.birth_date) { setValidation('La date d’entrée ne peut pas précéder la naissance.'); return }
    }
    if (!external && !nirEnabled && (socialSecurityNumber.trim() || clearSocialSecurityNumber)) { setValidation('La protection du numéro est indisponible. Effacez la nouvelle saisie ou annulez la suppression pour enregistrer les autres informations.'); return }
    setValidation('')
    const input: PersonnelWriteInput = { ...value, email: email.trim(), full_name: value.full_name.trim(), ...(clearSocialSecurityNumber ? { social_security_number: null } : socialSecurityNumber.trim() ? { social_security_number: socialSecurityNumber.trim() } : {}) }
    let accountResult: { welcome_email_sent: boolean } | undefined
    const ok = await action.run(async () => {
      if (external) {
        const result = await adminApi.create(token, { ...organization, employee_type: 'salarie', first_name: value.first_name.trim(), last_name: value.last_name.trim(), new_mail: email.trim(), generate_password: true })
        accountResult = { welcome_email_sent: result.welcome_email_sent === true }
        return result
      }
      if (employee) return personnelApi.update(token, employee.employee_id, input, version!)
      const { employee_type: _type, ...account } = organization
      const result = await personnelApi.create(token, { ...input, initial_period: period, ...(createAccount ? { account } : {}) })
      accountResult = result.account_creation
      return result
    })
    if (ok) {
      if (!employee || value.full_name.trim() !== employee.full_name || value.linked_username !== employee.linked_username || email.trim() !== (employee.email || '')) void client.invalidateQueries()
      setSocialSecurityNumber(''); setClearSocialSecurityNumber(false)
      if (accountResult) { setCreatedAccount(accountResult); onSaved(external ? 'Compte créé, sans inscription au liste des utilisateurs.' : 'Fiche et compte créés.'); return }
      onSaved(employee ? 'Fiche du personnel mise à jour.' : 'Personne enregistrée dans le registre.'); onClose()
    }
  }
  return <Dialog open fullWidth maxWidth="md" aria-labelledby="personnel-editor-title" onClose={() => { if (!action.busy) close(onClose) }}>
    <Box component="form" onSubmit={event => { event.preventDefault(); void save() }}>
      <DialogTitle id="personnel-editor-title">{employee ? `Modifier la fiche · ${employee.full_name}` : organization.role_tags.includes('SuperAdmin') ? 'Créer un compte système' : external ? 'Créer un compte externe' : 'Enregistrer une personne'}</DialogTitle>
      <DialogContent ref={completionRoot} dividers>
        {createdAccount ? <Alert severity={createdAccount.welcome_email_sent ? 'success' : 'warning'}>
          {external ? 'Compte créé sans inscription au liste des utilisateurs. ' : 'Fiche et compte créés et associés. '}
          {createdAccount.welcome_email_sent ? 'Le mot de passe temporaire a été transmis par e-mail.' : 'L’envoi de l’e-mail n’a pas été confirmé. Ne recréez pas le compte : utilisez « Mot de passe oublié » pour définir un nouveau mot de passe.'}
        </Alert> : <>
        {!employee && !organization.role_tags.includes('SuperAdmin') && <FormControlLabel sx={{ mb: 1 }} control={<Checkbox checked={external} disabled={action.busy} onChange={event => { setExternal(event.target.checked); setValidation(''); setSocialSecurityNumber(''); setNumberVisible(false) }} />} label={<Typography sx={{ fontWeight: 600 }}>Personne extérieure à l’entreprise</Typography>} />}
        <Typography variant="body2" color="text.secondary" sx={{ mb: 2 }}>{external ? 'Compte de connexion uniquement, sans fiche au liste des utilisateurs ni contrat.' : employee && !strictRegistration ? 'Ancienne fiche : complétez les champs manquants.' : 'Champs obligatoires : *.'}</Typography>
        {validation && <Alert severity="error" sx={{ mb: 2 }}>{validation}</Alert>}
        <WriteError error={action.error} conflict={action.conflict} />
        <Box sx={{ display: 'grid', gridTemplateColumns: { xs: '1fr', sm: '1fr 1fr' }, gap: 2 }}>
          {employee && (!employee.last_name || !employee.first_name) && <Typography variant="body2" color="text.secondary" sx={{ gridColumn: '1 / -1' }}>Nom historique : {employee.full_name}</Typography>}
          <TextField fullWidth required={strictRegistration} data-completion-field="last_name" label={personnelFieldLabels.last_name} autoComplete="family-name" value={value.last_name} disabled={action.busy} onChange={event => field('last_name', event.target.value)} slotProps={{ htmlInput: { maxLength: 100 } }} />
          <TextField fullWidth required={strictRegistration} data-completion-field="first_name" label={personnelFieldLabels.first_name} autoComplete="given-name" value={value.first_name} disabled={action.busy} onChange={event => field('first_name', event.target.value)} slotProps={{ htmlInput: { maxLength: 100 } }} />
          <TextField fullWidth data-completion-field="email" label="E-mail" type="email" required={wantsAccount} autoComplete="email" value={email} disabled={action.busy}
            onChange={event => setEmail(event.target.value)} onBlur={event => { if (event.target.validity.valid) setCheckedEmail(email.trim()) }}
            error={emailTaken} helperText={emailTaken ? 'Adresse déjà utilisée, y compris par un compte désactivé ou archivé.' : undefined}
            slotProps={{ htmlInput: { maxLength: 254 } }} sx={{ gridColumn: '1 / -1' }} />
          {!external && <>
          <TextField fullWidth select required={strictRegistration} data-completion-field="personnel_type" label={personnelFieldLabels.personnel_type} value={value.personnel_type} disabled={action.busy} onChange={event => field('personnel_type', event.target.value as PersonnelType)}>{personnelTypes.map(item => <MenuItem key={item.value} value={item.value}>{item.label}</MenuItem>)}</TextField>
          <TextField fullWidth required={strictRegistration} data-completion-field="nationality" label={personnelFieldLabels.nationality} value={value.nationality} disabled={action.busy} onChange={event => field('nationality', event.target.value)} slotProps={{ htmlInput: { maxLength: 100 } }} />
          <TextField fullWidth select required={strictRegistration} data-completion-field="sex" label={personnelFieldLabels.sex} value={value.sex} disabled={action.busy} onChange={event => field('sex', event.target.value)}><MenuItem value="">Non renseigné</MenuItem><MenuItem value="F">Femme</MenuItem><MenuItem value="M">Homme</MenuItem></TextField>
          <FrenchDateField fullWidth required={strictRegistration} completionField="birth_date" label={personnelFieldLabels.birth_date} value={value.birth_date || ''} disabled={action.busy} onChange={next => field('birth_date', next || null)} />
          <TextField fullWidth required={strictRegistration} data-completion-field="address" label={personnelFieldLabels.address} multiline minRows={2} value={value.address} disabled={action.busy} onChange={event => field('address', event.target.value)} slotProps={{ htmlInput: { maxLength: 1000 } }} sx={{ gridColumn: '1 / -1' }} />
          <TextField fullWidth required={strictRegistration} data-completion-field="job_title" label={personnelFieldLabels.job_title} value={value.job_title} disabled={action.busy} onChange={event => field('job_title', event.target.value)} slotProps={{ htmlInput: { maxLength: 200 } }} />
          <TextField fullWidth required={strictRegistration} data-completion-field="qualification" label={personnelFieldLabels.qualification} value={value.qualification} disabled={action.busy} onChange={event => field('qualification', event.target.value)} slotProps={{ htmlInput: { maxLength: 200 } }} />
          <Box sx={{ gridColumn: '1 / -1' }}>
            <SensitiveNumber token={token} employeeId={employee?.employee_id || 0} hasNumber={!!employee?.has_social_security_number && !clearSocialSecurityNumber} renderInput={secret => {
              const visible = socialSecurityNumber ? numberVisible : !!secret.value
              const toggle = () => { if (socialSecurityNumber) setNumberVisible(current => !current); else secret.toggle() }
              return <TextField fullWidth type={visible ? 'text' : 'password'} autoComplete="off" label="N° de sécurité sociale"
                value={socialSecurityNumber || secret.value} placeholder={employee?.has_social_security_number && !clearSocialSecurityNumber ? '•'.repeat(15) : ''}
                disabled={action.busy || !nirEnabled || clearSocialSecurityNumber}
                onChange={event => setSocialSecurityNumber(event.target.value)}
                slotProps={{ inputLabel: { shrink: true }, htmlInput: { maxLength: 64, spellCheck: false, autoCapitalize: 'none', 'data-lpignore': 'true', 'data-1p-ignore': 'true' }, input: {
                  endAdornment: <InputAdornment position="end"><Tooltip title={visible ? 'Masquer le numéro' : 'Afficher le numéro'}><span><IconButton
                    aria-label={visible ? 'Masquer le numéro' : 'Afficher le numéro'} aria-pressed={visible}
                    disabled={action.busy || !nirEnabled || clearSocialSecurityNumber || secret.busy || (!socialSecurityNumber && (!employee?.has_social_security_number || !secret.enabled))}
                    onMouseDown={event => event.preventDefault()} onClick={toggle} edge="end">{visible ? <VisibilityOffOutlined /> : <VisibilityOutlined />}</IconButton></span></Tooltip></InputAdornment>,
                } }} />
            }} />
            {!!socialSecurityNumber && <Button size="small" disabled={action.busy} onClick={() => setSocialSecurityNumber('')}>Effacer la nouvelle saisie</Button>}
            {employee?.has_social_security_number && <FormControlLabel control={<Checkbox checked={clearSocialSecurityNumber} disabled={action.busy || (!nirEnabled && !clearSocialSecurityNumber)} onChange={event => { setClearSocialSecurityNumber(event.target.checked); if (event.target.checked) setSocialSecurityNumber('') }} />} label={<Typography variant="body2">Effacer explicitement le numéro actuellement enregistré</Typography>} />}
          </Box>
          </>}
        </Box>
        {!employee && !external && <>
          <Divider sx={{ my: 2 }} />
          <Typography variant="subtitle2" sx={{ mb: 2 }}>Entrée et contrat</Typography>
          {value.linked_username && <RegistrationContract key={value.linked_username} token={token} username={value.linked_username} value={period} onChange={setPeriod} busy={action.busy} />}
          <EmploymentFields entryRequired value={period} onChange={setPeriod} busy={action.busy} />
        </>}
        {!employee && <Box sx={{ my: 2, p: 2, border: 1, borderColor: wantsAccount ? 'primary.main' : 'divider', borderRadius: 2, bgcolor: wantsAccount ? 'action.selected' : undefined }}>
          {external ? <Typography sx={{ fontSize: 17, fontWeight: 700, color: 'primary.main', mb: 1 }}>Créer le compte de connexion</Typography> :
          <FormControlLabel sx={{ m: 0, mb: createAccount ? 1 : 0 }} control={<Checkbox checked={createAccount} disabled={action.busy || !capabilities.canCreateAccounts} onChange={event => { setCreateAccount(event.target.checked); if (event.target.checked) { field('linked_username', null); setLinkedDisplayName('') } }} />} label={<Typography sx={{ fontSize: 17, fontWeight: 700, color: createAccount ? 'primary.main' : 'text.primary' }}>Créer le compte de connexion</Typography>} />
          }
          {wantsAccount && <><Typography variant="body2" color="text.secondary" sx={{ mb: 2 }}>Mot de passe temporaire envoyé par e-mail.</Typography>
            {capabilities.data?.environment === 'local-snapshot' && <Typography variant="caption" color="text.secondary" sx={{ display: 'block', mb: 2 }}>Test local : aucun e-mail réel n’est envoyé.</Typography>}
            <RegistrationAccountFields token={token} value={{ ...organization, employee_type: value.personnel_type || 'salarie' }} onChange={next => {
              if (next.role_tags.includes('SuperAdmin')) {
                setExternal(true); setCreateAccount(true); setValidation(''); setSocialSecurityNumber(''); setNumberVisible(false)
                setOrganization({ ...next, role_tags: ['SuperAdmin', 'Admin'], manager_id: null, group_ids: [], managed_group_ids: [] })
              } else setOrganization(next)
            }} busy={action.busy} allowSuperAdmin={allowSuperAdmin} /></>}
        </Box>}
        {employee && <Box sx={{ mt: 2 }}><AccountLink token={token} employeeId={employee.employee_id} value={value.linked_username} currentDisplayName={linkedDisplayName} busy={action.busy} onChange={(next, name) => { field('linked_username', next); setLinkedDisplayName(name) }} /></Box>}
        </>}
      </DialogContent>
      <DialogActions sx={{ px: 3, py: 2 }}><Button disabled={action.busy} onClick={() => createdAccount ? onClose() : close(onClose)}>{createdAccount ? 'Fermer' : 'Annuler'}</Button>{!createdAccount && <Button type="submit" variant="contained" disabled={action.busy || emailTaken || (checksLoginEmail && emailLookup.isFetching) || !value.full_name.trim() || action.conflict || (wantsAccount && !capabilities.canCreateAccounts)}>{action.busy ? 'Enregistrement…' : employee ? 'Enregistrer les modifications' : organization.role_tags.includes('SuperAdmin') ? 'Créer le compte système' : external ? 'Créer le compte externe' : createAccount ? 'Enregistrer et créer le compte' : 'Enregistrer au registre'}</Button>}</DialogActions>
    </Box>
  </Dialog>
}
