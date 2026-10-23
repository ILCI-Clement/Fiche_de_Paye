import { Box, Checkbox, FormControlLabel, MenuItem, TextField, Typography } from '@mui/material'
import { FrenchDateField } from '../components/FrenchDateField'
import { contractTypes, youthContractTypes } from './api'
import type { ContractType, EmploymentInput, YouthContractType } from './api'

type EmploymentProps = { value: EmploymentInput; onChange: (value: EmploymentInput) => void; busy: boolean }

function ContractFields({ value, onChange, busy }: EmploymentProps) {
  return <>
    <TextField select fullWidth required={value.is_temporary_worker} label="Type de contrat" value={value.contract_type} disabled={busy} onChange={event => onChange({ ...value, contract_type: event.target.value as ContractType, other_contract: event.target.value === 'autre' ? value.other_contract : '' })}>
      {contractTypes.map(item => <MenuItem key={item.value} value={item.value}>{item.label}</MenuItem>)}
    </TextField>
    {value.contract_type === 'autre' && <TextField fullWidth required={value.is_temporary_worker} label="Précisez le contrat" value={value.other_contract} disabled={busy} onChange={event => onChange({ ...value, other_contract: event.target.value })} slotProps={{ htmlInput: { maxLength: 200 } }} />}
  </>
}

export function SpecialEmploymentFields({ value, onChange, busy }: EmploymentProps) {
  const text = (key: 'work_permit_type' | 'work_permit_number' | 'temporary_company_name' | 'temporary_company_address', label: string, multiline = false) => <TextField fullWidth required label={label} value={value[key]} disabled={busy}
    multiline={multiline} minRows={multiline ? 2 : undefined} onChange={event => onChange({ ...value, [key]: event.target.value })}
    slotProps={{ htmlInput: { maxLength: multiline ? 1000 : 200 } }} />
  const section = { border: 1, borderColor: 'divider', borderRadius: 1.5, px: 2, py: 1, gridColumn: '1 / -1' }
  const fields = { display: 'grid', gridTemplateColumns: { xs: '1fr', sm: '1fr 1fr' }, gap: 2, pt: 1, pb: 1 }
  return <>
    <Typography variant="subtitle2" sx={{ gridColumn: '1 / -1' }}>Situations particulières</Typography>
    <Box sx={section}>
      <FormControlLabel control={<Checkbox checked={value.is_foreign_worker} disabled={busy} onChange={event => onChange({ ...value, is_foreign_worker: event.target.checked })} />} label="Travailleur étranger · autorisation de travail" />
      {value.is_foreign_worker && <Box sx={fields}>{text('work_permit_type', 'Type du document')}{text('work_permit_number', 'N° du document')}</Box>}
    </Box>
    <Box sx={section}>
      <FormControlLabel control={<Checkbox checked={value.is_temporary_worker} disabled={busy} onChange={event => onChange({ ...value, is_temporary_worker: event.target.checked })} />} label="Travailleur temporaire" />
      {value.is_temporary_worker && <Box sx={fields}><ContractFields value={value} onChange={onChange} busy={busy} />{text('temporary_company_name', 'Entreprise de travail temporaire')}<Box sx={{ gridColumn: '1 / -1' }}>{text('temporary_company_address', 'Adresse de l’entreprise', true)}</Box></Box>}
    </Box>
    <Box sx={section}>
      <FormControlLabel control={<Checkbox checked={value.is_young_worker} disabled={busy} onChange={event => onChange({ ...value, is_young_worker: event.target.checked })} />} label="Jeune travailleur / Apprenti" />
      {value.is_young_worker && <Box sx={{ py: 1 }}><TextField fullWidth select required label="Contrat spécifique" value={value.youth_contract_type} disabled={busy} onChange={event => onChange({ ...value, youth_contract_type: event.target.value as YouthContractType })}>
        {youthContractTypes.map(item => <MenuItem key={item.value} value={item.value}>{item.label}</MenuItem>)}
      </TextField></Box>}
    </Box>
  </>
}

export function EmploymentFields({ value, onChange, busy, entryRequired = false }: EmploymentProps & { entryRequired?: boolean }) {
  return <Box sx={{ display: 'grid', gridTemplateColumns: { xs: '1fr', sm: '1fr 1fr' }, gap: 2 }}>
    <FrenchDateField required={entryRequired} label="Date d’entrée" value={value.entry_date || ''} disabled={busy} fullWidth onChange={entry_date => onChange({ ...value, entry_date: entry_date || null })} />
    <FrenchDateField label="Date de sortie effective" value={value.actual_exit_date || ''} min={value.entry_date || undefined} disabled={busy} fullWidth onChange={actual_exit_date => onChange({ ...value, actual_exit_date: actual_exit_date || null })} />
    <FrenchDateField label="Fin prévue du contrat" value={value.exit_date || ''} min={value.entry_date || undefined} disabled={busy} fullWidth onChange={exit_date => onChange({ ...value, exit_date: exit_date || null })} />
    {!value.is_temporary_worker && <ContractFields value={value} onChange={onChange} busy={busy} />}
    <SpecialEmploymentFields value={value} onChange={onChange} busy={busy} />
  </Box>
}
