import { Box, Checkbox, FormControlLabel, MenuItem, TextField, Typography } from '@mui/material'
import { youthContractTypes } from './api'
import type { EmploymentInput, YouthContractType } from './api'
import { ContractFields } from './ContractFields'

export function SpecialEmploymentFields({ value, onChange, busy }: { value: EmploymentInput; onChange: (value: EmploymentInput) => void; busy: boolean }) {
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
