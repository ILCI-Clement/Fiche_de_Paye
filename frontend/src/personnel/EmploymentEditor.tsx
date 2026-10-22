import { useState } from 'react'
import { Alert, Box, Button, Dialog, DialogActions, DialogContent, DialogTitle, Typography } from '@mui/material'
import { useDraftGuard } from '../draftGuard'
import { blankPeriod, personnelApi } from './api'
import type { Employee, EmploymentPeriod } from './api'
import { EmploymentFields } from './EmploymentFields'
import { WriteError } from './WriteError'
import { useRegistryAction } from './useRegistryAction'

export function EmploymentEditor({ token, employee, period, onClose, onSaved }: { token: string; employee: Employee; period?: EmploymentPeriod; onClose: () => void; onSaved: (message: string) => void }) {
  const [version] = useState(employee.version)
  const initial = () => period ? { entry_date: period.entry_date, exit_date: period.exit_date, contract_type: period.contract_type, other_contract: period.other_contract } : blankPeriod()
  const [value, setValue] = useState(initial)
  const [baseline] = useState(() => JSON.stringify(initial()))
  const [validation, setValidation] = useState('')
  const action = useRegistryAction()
  const close = useDraftGuard(JSON.stringify(value) !== baseline, action.busy, 'Période d’emploi')
  async function save() {
    if (action.busy) return
    if (value.entry_date && value.exit_date && value.exit_date < value.entry_date) { setValidation('La date de sortie doit être postérieure ou égale à la date d’entrée.'); return }
    setValidation('')
    const ok = await action.run(() => personnelApi.savePeriod(token, employee.employee_id, period?.period_id ?? null, value, version))
    if (ok) { onSaved(period ? 'Période d’emploi corrigée ; la version précédente reste dans l’historique.' : 'Nouvelle période d’emploi enregistrée.'); onClose() }
  }
  return <Dialog open fullWidth maxWidth="sm" aria-labelledby="employment-editor-title" onClose={() => { if (!action.busy) close(onClose) }}>
    <Box component="form" onSubmit={event => { event.preventDefault(); void save() }}>
      <DialogTitle id="employment-editor-title">{period ? 'Corriger une période d’emploi' : 'Ajouter une période d’emploi'}</DialogTitle>
      <DialogContent dividers>
        <Typography sx={{ mb: 1, fontWeight: 500 }}>{employee.full_name}</Typography>
        <Typography variant="body2" color="text.secondary" sx={{ mb: 2 }}>Dates et contrat sont facultatifs. Une nouvelle entrée après un départ est enregistrée comme une nouvelle période. L’état du compte de connexion reste inchangé.</Typography>
        {validation && <Alert severity="error" sx={{ mb: 2 }}>{validation}</Alert>}
        <WriteError error={action.error} conflict={action.conflict} />
        <EmploymentFields token={token} username={employee.linked_username} accountDisplayName={employee.linked_display_name} value={value} onChange={setValue} busy={action.busy} />
      </DialogContent>
      <DialogActions sx={{ p: 2 }}><Button disabled={action.busy} onClick={() => close(onClose)}>Annuler</Button><Button type="submit" variant="contained" disabled={action.busy || action.conflict}>{action.busy ? 'Enregistrement…' : 'Enregistrer la période'}</Button></DialogActions>
    </Box>
  </Dialog>
}
