import { useState } from 'react'
import { Alert, Box, Button, Dialog, DialogActions, DialogContent, DialogTitle, Typography } from '@mui/material'
import { useDraftGuard } from '../draftGuard'
import { blankPeriod, dateLabel, personnelApi } from './api'
import type { Employee, EmploymentPeriod } from './api'
import { EmploymentFields, SpecialEmploymentFields } from './EmploymentFields'
import { WriteError } from './WriteError'
import { useRegistryAction } from './useRegistryAction'
import { validateEmployment } from './personnelValidation'

export function EmploymentEditor({ token, employee, period, situationsOnly = false, onClose, onSaved }: { token: string; employee: Employee; period?: EmploymentPeriod; situationsOnly?: boolean; onClose: () => void; onSaved: (message: string) => void }) {
  const [version] = useState(employee.version)
  const initial = () => {
    const value = blankPeriod()
    if (period) for (const key of Object.keys(value) as (keyof typeof value)[]) Object.assign(value, { [key]: period[key] ?? value[key] })
    return value
  }
  const [value, setValue] = useState(initial)
  const [baseline] = useState(() => JSON.stringify(initial()))
  const [validation, setValidation] = useState('')
  const action = useRegistryAction()
  const editSituations = situationsOnly && !!period
  const close = useDraftGuard(JSON.stringify(value) !== baseline, action.busy, editSituations ? 'Situations particulières' : 'Période d’emploi')
  async function save() {
    if (action.busy) return
    if ((!period || employee.registration_schema_version >= 2) && !value.entry_date) { setValidation('La date d’entrée est obligatoire.'); return }
    const problem = validateEmployment(value)
    if (problem) { setValidation(problem); return }
    setValidation('')
    const ok = await action.run(() => personnelApi.savePeriod(token, employee.employee_id, period?.period_id ?? null, value, version))
    if (ok) { onSaved(editSituations ? 'Situations particulières mises à jour.' : period ? 'Période d’emploi corrigée ; la version précédente reste dans l’historique.' : 'Nouvelle période d’emploi enregistrée.'); onClose() }
  }
  return <Dialog open fullWidth maxWidth="md" aria-labelledby="employment-editor-title" onClose={() => { if (!action.busy) close(onClose) }}>
    <Box component="form" onSubmit={event => { event.preventDefault(); void save() }}>
      <DialogTitle id="employment-editor-title">{editSituations ? 'Modifier les situations particulières' : period ? 'Corriger une période d’emploi' : 'Ajouter une période d’emploi'}</DialogTitle>
      <DialogContent dividers>
        <Typography sx={{ mb: 1, fontWeight: 500 }}>{employee.full_name}</Typography>
        <Typography variant="body2" color="text.secondary" sx={{ mb: 2 }}>{editSituations ? `Période d’emploi · entrée le ${dateLabel(value.entry_date)}` : `La sortie effective est indépendante de la fin prévue du contrat.${employee.linked_username ? ' Le début et la fin prévue restent synchronisés avec le contrat associé ; les dates incomplètes ne créent pas de contrat.' : ' Une nouvelle entrée après un départ est enregistrée comme une nouvelle période.'}`}</Typography>
        {validation && <Alert severity="error" sx={{ mb: 2 }}>{validation}</Alert>}
        <WriteError error={action.error} conflict={action.conflict} />
        {editSituations ? <Box sx={{ display: 'grid', gridTemplateColumns: 'minmax(0, 1fr)', gap: 2 }}><SpecialEmploymentFields value={value} onChange={setValue} busy={action.busy} /></Box> : <EmploymentFields entryRequired={!period || employee.registration_schema_version >= 2} value={value} onChange={setValue} busy={action.busy} />}
      </DialogContent>
      <DialogActions sx={{ p: 2 }}><Button disabled={action.busy} onClick={() => close(onClose)}>Annuler</Button><Button type="submit" variant="contained" disabled={action.busy || action.conflict}>{action.busy ? 'Enregistrement…' : 'Enregistrer la période'}</Button></DialogActions>
    </Box>
  </Dialog>
}
