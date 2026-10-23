import { Alert, Box, MenuItem, TextField } from '@mui/material'
import { useQuery } from '@tanstack/react-query'
import { adminApi } from '../admin/api'
import { dateLabel } from './api'
import type { EmploymentInput } from './api'
import { useRegistryViewer } from './context'

/** Explicit prefill: never silently replace an entry date already typed by HR. */
export function RegistrationContract({ token, username, value, onChange, busy }: { token: string; username: string; value: EmploymentInput; onChange: (value: EmploymentInput) => void; busy: boolean }) {
  const viewer = useRegistryViewer()
  const contracts = useQuery({ queryKey: ['registration-contracts', viewer, username], queryFn: ({ signal }) => adminApi.contracts(token, username, signal), retry: false, gcTime: 0 })
  const options = contracts.data?.contracts.filter(contract => !contract.is_archived) || []
  const selected = options.find(contract => contract.start_date === value.entry_date && contract.end_date === value.exit_date)
  return <Box sx={{ mb: 2 }}>
    {contracts.error ? <Alert severity="warning">Les contrats du compte ne peuvent pas être chargés. L’inscription reste possible ; vérifiez les dates dans Administration.</Alert> : !!options.length && <TextField fullWidth select label="Préremplir depuis un contrat du compte" value={selected?.id || ''} disabled={busy || contracts.isPending}
      onChange={event => { const contract = options.find(item => item.id === Number(event.target.value)); if (contract) onChange({ ...value, entry_date: contract.start_date, exit_date: contract.end_date, contract_type: contract.is_cdi ? 'CDI' : contract.employee_type === 'stagiaire' ? 'stage' : value.contract_type }) }}>
      <MenuItem value="">Saisie manuelle</MenuItem>
      {options.map(contract => <MenuItem key={contract.id} value={contract.id}>{dateLabel(contract.start_date)} – {contract.is_cdi ? 'CDI' : dateLabel(contract.end_date)}</MenuItem>)}
    </TextField>}
  </Box>
}
