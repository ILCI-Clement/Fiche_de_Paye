import { Box } from '@mui/material'
import { FrenchDateField } from '../components/FrenchDateField'
import type { EmploymentInput } from './api'
import { SpecialEmploymentFields } from './SpecialEmploymentFields'
import { ContractFields } from './ContractFields'

export function EmploymentFields({ value, onChange, busy, entryRequired = false }: { value: EmploymentInput; onChange: (value: EmploymentInput) => void; busy: boolean; entryRequired?: boolean }) {
  return <Box sx={{ display: 'grid', gridTemplateColumns: { xs: '1fr', sm: '1fr 1fr' }, gap: 2 }}>
    <FrenchDateField required={entryRequired} label="Date d’entrée" value={value.entry_date || ''} disabled={busy} fullWidth onChange={entry_date => onChange({ ...value, entry_date: entry_date || null })} />
    <FrenchDateField label="Date de sortie effective" value={value.actual_exit_date || ''} min={value.entry_date || undefined} disabled={busy} fullWidth onChange={actual_exit_date => onChange({ ...value, actual_exit_date: actual_exit_date || null })} />
    <FrenchDateField label="Fin prévue du contrat" value={value.exit_date || ''} min={value.entry_date || undefined} disabled={busy} fullWidth onChange={exit_date => onChange({ ...value, exit_date: exit_date || null })} />
    {!value.is_temporary_worker && <ContractFields value={value} onChange={onChange} busy={busy} />}
    <SpecialEmploymentFields value={value} onChange={onChange} busy={busy} />
  </Box>
}
