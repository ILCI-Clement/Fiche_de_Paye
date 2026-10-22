import { Box, MenuItem, TextField } from '@mui/material'
import { FrenchDateField } from '../components/FrenchDateField'
import { contractTypes } from './api'
import type { ContractType, EmploymentInput } from './api'
import { ContractPrefill } from './ContractPrefill'

export function EmploymentFields({ value, onChange, busy, token, username, accountDisplayName }: { value: EmploymentInput; onChange: (value: EmploymentInput) => void; busy: boolean; token: string; username: string | null; accountDisplayName?: string | null }) {
  return <Box sx={{ display: 'grid', gridTemplateColumns: { xs: '1fr', sm: '1fr 1fr' }, gap: 2 }}>
    <FrenchDateField label="Date d’entrée" value={value.entry_date || ''} disabled={busy} fullWidth onChange={entry_date => onChange({ ...value, entry_date: entry_date || null })} />
    <FrenchDateField label="Date de sortie" value={value.exit_date || ''} min={value.entry_date || undefined} disabled={busy} fullWidth onChange={exit_date => onChange({ ...value, exit_date: exit_date || null })} />
    <TextField select fullWidth label="Type de contrat" value={value.contract_type} disabled={busy} onChange={event => onChange({ ...value, contract_type: event.target.value as ContractType, other_contract: event.target.value === 'autre' ? value.other_contract : '' })}>
      {contractTypes.map(item => <MenuItem key={item.value} value={item.value}>{item.label}</MenuItem>)}
    </TextField>
    {value.contract_type === 'autre' && <TextField fullWidth label="Précisez le contrat" value={value.other_contract} disabled={busy} onChange={event => onChange({ ...value, other_contract: event.target.value })} slotProps={{ htmlInput: { maxLength: 200 } }} />}
    <ContractPrefill key={username || 'unlinked'} token={token} username={username} accountDisplayName={accountDisplayName} busy={busy} onApply={update => onChange({ ...value, ...update })} />
  </Box>
}
