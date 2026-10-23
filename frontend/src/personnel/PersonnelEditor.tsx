import { useState } from 'react'
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
import { useQueryClient } from '@tanstack/react-query'

function initialInput(employee?: Employee): PersonnelInput {
  return { full_name: employee?.full_name || '', last_name: employee?.last_name || '', first_name: employee?.first_name || '', nationality: employee?.nationality || '', sex: employee?.sex || '', address: employee?.address || '', birth_date: employee?.birth_date || null, job_title: employee?.job_title || '', qualification: employee?.qualification || '', personnel_type: employee?.personnel_type || '', linked_username: employee?.linked_username || null }
}

export function PersonnelEditor({ token, employee, onClose, onSaved }: { token: string; employee?: Employee; onClose: () => void; onSaved: (message: string) => void }) {
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
  const capability = useSensitiveNumberCapability(token)
  const nirEnabled = capability.data?.nir_enabled === true && !capability.error
  const action = useRegistryAction()
  const strictRegistration = !employee || employee.registration_schema_version >= 2
  const dirty = JSON.stringify(value) !== initial || JSON.stringify(period) !== JSON.stringify(blankPeriod()) || !!socialSecurityNumber || clearSocialSecurityNumber
  const close = useDraftGuard(dirty, action.busy, employee ? 'Fiche du personnel' : 'Nouvelle inscription au registre')
  const field = (key: keyof PersonnelInput, next: string | null) => setValue(current => {
    const updated = { ...current, [key]: next }
    if (updated.first_name.trim() && updated.last_name.trim()) updated.full_name = `${updated.first_name.trim()} ${updated.last_name.trim()}`
    return updated
  })
  async function save() {
    if (action.busy) return
    if (!value.full_name.trim()) { setValidation('Renseignez le nom et les prénoms.'); return }
    if (strictRegistration) {
      const missing = missingRegistrationFields(value, employee ? { ...blankPeriod(), entry_date: employee.latest_period?.entry_date || null } : period)
      if (missing.length) { setValidation(`Champs obligatoires : ${missing.join(', ')}.`); return }
    }
    if (!employee) {
      const periodError = validateEmployment(period)
      if (periodError) { setValidation(periodError); return }
      if (value.birth_date && period.entry_date! < value.birth_date) { setValidation('La date d’entrée ne peut pas précéder la naissance.'); return }
    }
    if (!nirEnabled && (socialSecurityNumber.trim() || clearSocialSecurityNumber)) { setValidation('La protection du numéro est indisponible. Effacez la nouvelle saisie ou annulez la suppression pour enregistrer les autres informations.'); return }
    setValidation('')
    const input: PersonnelWriteInput = { ...value, full_name: value.full_name.trim(), ...(clearSocialSecurityNumber ? { social_security_number: null } : socialSecurityNumber.trim() ? { social_security_number: socialSecurityNumber.trim() } : {}) }
    const ok = await action.run(() => employee ? personnelApi.update(token, employee.employee_id, input, version!) : personnelApi.create(token, { ...input, initial_period: period }))
    if (ok) {
      if (!employee || value.full_name.trim() !== employee.full_name || value.linked_username !== employee.linked_username) void client.invalidateQueries()
      setSocialSecurityNumber(''); setClearSocialSecurityNumber(false); onSaved(employee ? 'Fiche du personnel mise à jour.' : 'Personne enregistrée dans le registre.'); onClose()
    }
  }
  return <Dialog open fullWidth maxWidth="md" aria-labelledby="personnel-editor-title" onClose={() => { if (!action.busy) close(onClose) }}>
    <Box component="form" onSubmit={event => { event.preventDefault(); void save() }}>
      <DialogTitle id="personnel-editor-title">{employee ? `Modifier la fiche · ${employee.full_name}` : 'Enregistrer un salarié'}</DialogTitle>
      <DialogContent dividers>
        <Typography variant="body2" color="text.secondary" sx={{ mb: 2 }}>{employee && !strictRegistration ? 'Les anciennes fiches peuvent être complétées progressivement. Chaque correction est conservée dans l’historique.' : 'Les champs marqués * sont obligatoires. Le numéro de sécurité sociale et le compte de connexion sont facultatifs.'}</Typography>
        {validation && <Alert severity="error" sx={{ mb: 2 }}>{validation}</Alert>}
        <WriteError error={action.error} conflict={action.conflict} />
        <Box sx={{ display: 'grid', gridTemplateColumns: { xs: '1fr', sm: '1fr 1fr' }, gap: 2 }}>
          {employee && (!employee.last_name || !employee.first_name) && <Typography variant="body2" color="text.secondary" sx={{ gridColumn: '1 / -1' }}>Nom historique : {employee.full_name}</Typography>}
          <TextField fullWidth required={strictRegistration} label={personnelFieldLabels.last_name} autoComplete="family-name" value={value.last_name} disabled={action.busy} onChange={event => field('last_name', event.target.value)} slotProps={{ htmlInput: { maxLength: 100 } }} />
          <TextField fullWidth required={strictRegistration} label={personnelFieldLabels.first_name} autoComplete="given-name" value={value.first_name} disabled={action.busy} onChange={event => field('first_name', event.target.value)} slotProps={{ htmlInput: { maxLength: 100 } }} />
          <TextField fullWidth select required={strictRegistration} label={personnelFieldLabels.personnel_type} value={value.personnel_type} disabled={action.busy} onChange={event => field('personnel_type', event.target.value as PersonnelType)}>{personnelTypes.map(item => <MenuItem key={item.value} value={item.value}>{item.label}</MenuItem>)}</TextField>
          <TextField fullWidth required={strictRegistration} label={personnelFieldLabels.nationality} value={value.nationality} disabled={action.busy} onChange={event => field('nationality', event.target.value)} slotProps={{ htmlInput: { maxLength: 100 } }} />
          <TextField fullWidth select required={strictRegistration} label={personnelFieldLabels.sex} value={value.sex} disabled={action.busy} onChange={event => field('sex', event.target.value)}><MenuItem value="">Non renseigné</MenuItem><MenuItem value="F">Femme</MenuItem><MenuItem value="M">Homme</MenuItem></TextField>
          <FrenchDateField fullWidth required={strictRegistration} label={personnelFieldLabels.birth_date} value={value.birth_date || ''} disabled={action.busy} onChange={next => field('birth_date', next || null)} />
          <TextField fullWidth required={strictRegistration} label={personnelFieldLabels.address} multiline minRows={2} value={value.address} disabled={action.busy} onChange={event => field('address', event.target.value)} slotProps={{ htmlInput: { maxLength: 1000 } }} sx={{ gridColumn: '1 / -1' }} />
          <TextField fullWidth required={strictRegistration} label={personnelFieldLabels.job_title} value={value.job_title} disabled={action.busy} onChange={event => field('job_title', event.target.value)} slotProps={{ htmlInput: { maxLength: 200 } }} />
          <TextField fullWidth required={strictRegistration} label={personnelFieldLabels.qualification} value={value.qualification} disabled={action.busy} onChange={event => field('qualification', event.target.value)} slotProps={{ htmlInput: { maxLength: 200 } }} />
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
        </Box>
        {!employee && <>
          <Divider sx={{ my: 2 }} />
          <Typography variant="subtitle2" sx={{ mb: 2 }}>Entrée et contrat</Typography>
          {value.linked_username && <RegistrationContract key={value.linked_username} token={token} username={value.linked_username} value={period} onChange={setPeriod} busy={action.busy} />}
          <EmploymentFields entryRequired value={period} onChange={setPeriod} busy={action.busy} />
        </>}
        <Box sx={{ mt: 2 }}><AccountLink token={token} employeeId={employee?.employee_id} value={value.linked_username} currentDisplayName={linkedDisplayName} busy={action.busy} onChange={(next, name) => { field('linked_username', next); setLinkedDisplayName(name) }} /></Box>
      </DialogContent>
      <DialogActions sx={{ px: 3, py: 2 }}><Button disabled={action.busy} onClick={() => close(onClose)}>Annuler</Button><Button type="submit" variant="contained" disabled={action.busy || !value.full_name.trim() || action.conflict}>{action.busy ? 'Enregistrement…' : employee ? 'Enregistrer les modifications' : 'Enregistrer au registre'}</Button></DialogActions>
    </Box>
  </Dialog>
}
