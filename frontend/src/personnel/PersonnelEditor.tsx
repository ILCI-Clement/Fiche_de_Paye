import { useState } from 'react'
import { Alert, Box, Button, Checkbox, Dialog, DialogActions, DialogContent, DialogTitle, Divider, FormControlLabel, MenuItem, TextField, Typography } from '@mui/material'
import { useDraftGuard } from '../draftGuard'
import { FrenchDateField } from '../components/FrenchDateField'
import { personnelApi, personnelTypes, blankPeriod } from './api'
import type { Employee, PersonnelInput, PersonnelType, PersonnelWriteInput } from './api'
import { AccountLink } from './AccountLink'
import { EmploymentFields } from './EmploymentFields'
import { useRegistryAction } from './useRegistryAction'
import { WriteError } from './WriteError'
import { useSensitiveNumberCapability } from './useSensitiveNumberCapability'
import { useQueryClient } from '@tanstack/react-query'

function initialInput(employee?: Employee): PersonnelInput {
  return { full_name: employee?.full_name || '', nationality: employee?.nationality || '', sex: employee?.sex || '', address: employee?.address || '', birth_date: employee?.birth_date || null, job_title: employee?.job_title || '', qualification: employee?.qualification || '', personnel_type: employee?.personnel_type || '', linked_username: employee?.linked_username || null }
}

export function PersonnelEditor({ token, employee, onClose, onSaved }: { token: string; employee?: Employee; onClose: () => void; onSaved: (message: string) => void }) {
  const client = useQueryClient()
  const [version] = useState(employee?.version)
  const [value, setValue] = useState(() => initialInput(employee))
  const [linkedDisplayName, setLinkedDisplayName] = useState(employee?.linked_display_name || '')
  const [initial] = useState(() => JSON.stringify(initialInput(employee)))
  const [withPeriod, setWithPeriod] = useState(false)
  const [period, setPeriod] = useState(blankPeriod)
  const [validation, setValidation] = useState('')
  const [socialSecurityNumber, setSocialSecurityNumber] = useState('')
  const [clearSocialSecurityNumber, setClearSocialSecurityNumber] = useState(false)
  const capability = useSensitiveNumberCapability(token)
  const nirEnabled = capability.data?.nir_enabled === true && !capability.error
  const action = useRegistryAction()
  const dirty = JSON.stringify(value) !== initial || withPeriod || JSON.stringify(period) !== JSON.stringify(blankPeriod()) || !!socialSecurityNumber || clearSocialSecurityNumber
  const close = useDraftGuard(dirty, action.busy, employee ? 'Fiche du personnel' : 'Nouvelle inscription au registre')
  const field = (key: keyof PersonnelInput, next: string | null) => setValue(current => ({ ...current, [key]: next }))
  async function save() {
    if (action.busy) return
    if (!value.full_name.trim()) { setValidation('Renseignez le nom et les prénoms.'); return }
    if (withPeriod && period.entry_date && period.exit_date && period.exit_date < period.entry_date) { setValidation('La date de sortie doit être postérieure ou égale à la date d’entrée.'); return }
    if (!nirEnabled && (socialSecurityNumber.trim() || clearSocialSecurityNumber)) { setValidation('La protection du numéro est indisponible. Effacez la nouvelle saisie ou annulez la suppression pour enregistrer les autres informations.'); return }
    setValidation('')
    const input: PersonnelWriteInput = { ...value, full_name: value.full_name.trim(), ...(clearSocialSecurityNumber ? { social_security_number: null } : socialSecurityNumber.trim() ? { social_security_number: socialSecurityNumber.trim() } : {}) }
    const ok = await action.run(() => employee ? personnelApi.update(token, employee.employee_id, input, version!) : personnelApi.create(token, { ...input, ...(withPeriod ? { initial_period: period } : {}) }))
    if (ok) {
      if (!employee || value.full_name.trim() !== employee.full_name || value.linked_username !== employee.linked_username) void client.invalidateQueries()
      setSocialSecurityNumber(''); setClearSocialSecurityNumber(false); onSaved(employee ? 'Fiche du personnel mise à jour.' : 'Personne enregistrée dans le registre.'); onClose()
    }
  }
  return <Dialog open fullWidth maxWidth="md" aria-labelledby="personnel-editor-title" onClose={() => { if (!action.busy) close(onClose) }}>
    <Box component="form" onSubmit={event => { event.preventDefault(); void save() }}>
      <DialogTitle id="personnel-editor-title">{employee ? `Modifier la fiche · ${employee.full_name}` : 'Enregistrer un salarié'}</DialogTitle>
      <DialogContent dividers>
        <Typography variant="body2" color="text.secondary" sx={{ mb: 2 }}>Seuls le nom et les prénoms sont obligatoires. {employee ? 'Chaque correction est conservée dans l’historique.' : 'Le numéro de registre est attribué automatiquement. L’inscription ne crée pas de compte de connexion.'}</Typography>
        {validation && <Alert severity="error" sx={{ mb: 2 }}>{validation}</Alert>}
        <WriteError error={action.error} conflict={action.conflict} />
        <Box sx={{ display: 'grid', gridTemplateColumns: { xs: '1fr', sm: '1fr 1fr' }, gap: 2 }}>
          <TextField fullWidth required label="Nom et prénoms" autoComplete="off" value={value.full_name} disabled={action.busy} onChange={event => field('full_name', event.target.value)} slotProps={{ htmlInput: { maxLength: 200 } }} sx={{ gridColumn: '1 / -1' }} />
          <TextField fullWidth select label="Type de personnel" value={value.personnel_type} disabled={action.busy} onChange={event => field('personnel_type', event.target.value as PersonnelType)}>{personnelTypes.map(item => <MenuItem key={item.value} value={item.value}>{item.label}</MenuItem>)}</TextField>
          <TextField fullWidth label="Nationalité" value={value.nationality} disabled={action.busy} onChange={event => field('nationality', event.target.value)} slotProps={{ htmlInput: { maxLength: 100 } }} />
          <TextField fullWidth select label="Sexe" value={value.sex} disabled={action.busy} onChange={event => field('sex', event.target.value)}><MenuItem value="">Non renseigné</MenuItem><MenuItem value="F">Femme</MenuItem><MenuItem value="M">Homme</MenuItem></TextField>
          <FrenchDateField fullWidth label="Date de naissance" value={value.birth_date || ''} disabled={action.busy} onChange={next => field('birth_date', next || null)} />
          <TextField fullWidth label="Adresse" multiline minRows={2} value={value.address} disabled={action.busy} onChange={event => field('address', event.target.value)} slotProps={{ htmlInput: { maxLength: 1000 } }} sx={{ gridColumn: '1 / -1' }} />
          <TextField fullWidth label="Emploi" value={value.job_title} disabled={action.busy} onChange={event => field('job_title', event.target.value)} slotProps={{ htmlInput: { maxLength: 200 } }} />
          <TextField fullWidth label="Qualification" value={value.qualification} disabled={action.busy} onChange={event => field('qualification', event.target.value)} slotProps={{ htmlInput: { maxLength: 200 } }} />
          <Box sx={{ gridColumn: '1 / -1' }}>
            <TextField fullWidth type="password" autoComplete="off" label={employee?.has_social_security_number ? 'Remplacer le n° de sécurité sociale' : 'N° de sécurité sociale'} value={socialSecurityNumber} disabled={action.busy || !nirEnabled || clearSocialSecurityNumber} onChange={event => setSocialSecurityNumber(event.target.value)}
              slotProps={{ htmlInput: { maxLength: 64, spellCheck: false, autoCapitalize: 'none', 'data-lpignore': 'true', 'data-1p-ignore': 'true' } }}
              helperText={capability.isPending ? 'Vérification de la protection…' : !nirEnabled ? 'Saisie protégée indisponible dans cet environnement. Les autres informations restent modifiables.' : employee?.has_social_security_number ? 'Un numéro est déjà enregistré et chiffré. Laissez ce champ vide pour le conserver ; aucune valeur existante n’est préchargée.' : 'Facultatif. Le numéro sera chiffré avant stockage ; il n’apparaît pas dans l’historique.'} />
            {!!socialSecurityNumber && <Button size="small" disabled={action.busy} onClick={() => setSocialSecurityNumber('')}>Effacer la nouvelle saisie</Button>}
            {employee?.has_social_security_number && <FormControlLabel control={<Checkbox checked={clearSocialSecurityNumber} disabled={action.busy || (!nirEnabled && !clearSocialSecurityNumber)} onChange={event => { setClearSocialSecurityNumber(event.target.checked); if (event.target.checked) setSocialSecurityNumber('') }} />} label={<Typography variant="body2">Effacer explicitement le numéro actuellement enregistré</Typography>} />}
          </Box>
          <AccountLink token={token} employeeId={employee?.employee_id} value={value.linked_username} currentDisplayName={linkedDisplayName} busy={action.busy} onChange={(next, name) => { field('linked_username', next); setLinkedDisplayName(name) }} />
        </Box>
        {!employee && <>
          <Divider sx={{ my: 2 }} />
          <FormControlLabel control={<Checkbox checked={withPeriod} disabled={action.busy} onChange={event => setWithPeriod(event.target.checked)} />} label="Ajouter une première période d’emploi (facultatif)" />
          {withPeriod && <Box sx={{ mt: 2 }}><EmploymentFields token={token} username={value.linked_username} accountDisplayName={linkedDisplayName} value={period} onChange={setPeriod} busy={action.busy} /></Box>}
        </>}
      </DialogContent>
      <DialogActions sx={{ px: 3, py: 2 }}><Button disabled={action.busy} onClick={() => close(onClose)}>Annuler</Button><Button type="submit" variant="contained" disabled={action.busy || !value.full_name.trim() || action.conflict}>{action.busy ? 'Enregistrement…' : employee ? 'Enregistrer les modifications' : 'Enregistrer au registre'}</Button></DialogActions>
    </Box>
  </Dialog>
}
