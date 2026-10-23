import type { ReactNode } from 'react'
import { Box, Typography } from '@mui/material'
import { dateLabel, typeLabel, youthContractLabel } from './api'
import type { EmploymentInput, PersonnelInput } from './api'
import { personnelFieldLabels } from './personnelValidation'

/** Shared read-only fields; sensitive values remain explicit, permission-controlled children. */
export function PersonnelInformation({ employee, structuredNames = false, accountName, children }: {
  employee: PersonnelInput; structuredNames?: boolean; accountName?: string; children?: ReactNode;
}) {
  const keys = (Object.keys(personnelFieldLabels) as (keyof typeof personnelFieldLabels)[])
    .filter(key => structuredNames || (key !== 'last_name' && key !== 'first_name'))
  return <Box component="dl" sx={{ m: 0, display: 'grid', gridTemplateColumns: { xs: '1fr', sm: '1fr 1fr' }, gap: 2.5 }}>
    {keys.map(key => {
      const value = key === 'personnel_type' ? typeLabel(employee.personnel_type) : key === 'birth_date' ? dateLabel(employee.birth_date)
        : key === 'sex' ? employee.sex === 'F' ? 'Femme' : employee.sex === 'M' ? 'Homme' : '' : employee[key]
      return <InformationValue key={key} label={personnelFieldLabels[key]} value={value} wide={key === 'address'} />
    })}
    {accountName !== undefined && <InformationValue label="Compte associé" value={accountName} />}
    {children}
  </Box>
}

function InformationValue({ label, value, wide }: { label: string; value: string | null; wide?: boolean }) {
  return <Box component="div" sx={wide ? { gridColumn: '1 / -1' } : undefined}>
    <Typography component="dt" variant="caption" color="text.secondary">{label}</Typography>
    <Typography component="dd" variant="body2" sx={{ m: 0, mt: .5, whiteSpace: 'pre-wrap', overflowWrap: 'anywhere' }}>{value || 'Non renseigné'}</Typography>
  </Box>
}

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
