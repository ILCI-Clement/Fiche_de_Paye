import { MenuItem, TextField } from '@mui/material'
import { contractTypes } from './api'
import type { ContractType, EmploymentInput } from './api'

export function ContractFields({ value, onChange, busy }: { value: EmploymentInput; onChange: (value: EmploymentInput) => void; busy: boolean }) {
  return <>
    <TextField select fullWidth required={value.is_temporary_worker} label="Type de contrat" value={value.contract_type} disabled={busy} onChange={event => onChange({ ...value, contract_type: event.target.value as ContractType, other_contract: event.target.value === 'autre' ? value.other_contract : '' })}>
      {contractTypes.map(item => <MenuItem key={item.value} value={item.value}>{item.label}</MenuItem>)}
    </TextField>
    {value.contract_type === 'autre' && <TextField fullWidth required={value.is_temporary_worker} label="Précisez le contrat" value={value.other_contract} disabled={busy} onChange={event => onChange({ ...value, other_contract: event.target.value })} slotProps={{ htmlInput: { maxLength: 200 } }} />}
  </>
}
