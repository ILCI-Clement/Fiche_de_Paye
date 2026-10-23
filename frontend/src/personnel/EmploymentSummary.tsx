import { Box, Typography } from '@mui/material'
import { youthContractLabel } from './api'
import type { EmploymentInput } from './api'

export function EmploymentSummary({ period }: { period: EmploymentInput }) {
  const rows = [
    ...(period.is_foreign_worker ? [['Autorisation de travail', `${period.work_permit_type} · ${period.work_permit_number}`]] : []),
    ...(period.is_temporary_worker ? [['Entreprise de travail temporaire', `${period.temporary_company_name}\n${period.temporary_company_address}`]] : []),
    ...(period.is_young_worker ? [['Jeune travailleur / Apprenti', youthContractLabel(period.youth_contract_type)]] : []),
  ]
  return !!rows.length && <Box sx={{ mt: 1, display: 'grid', gap: .75 }}>{rows.map(([label, value]) => <Box key={label}>
    <Typography variant="caption" color="text.secondary">{label}</Typography>
    <Typography variant="body2" sx={{ whiteSpace: 'pre-line', overflowWrap: 'anywhere' }}>{value}</Typography>
  </Box>)}</Box>
}
