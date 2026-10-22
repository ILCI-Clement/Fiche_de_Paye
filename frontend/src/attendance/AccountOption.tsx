import { personDisplayName } from '../personDisplayName'
import { Box, Typography } from '@mui/material'
import { DepartmentBadges } from '../components/DepartmentBadge'
import type { Account } from './types'

export function AccountOption({ account, compact = false }: { account: Account; compact?: boolean }) {
  const name = personDisplayName(account)
  return <Box component="span" className={`account-option${compact ? ' account-option-compact' : ''}`}>
    <DepartmentBadges names={account.group_names || []} />
    <Box component="span" sx={{ minWidth: 0 }}>
      <Typography component="span" noWrap sx={{ display: 'block', fontSize: 14, fontWeight: 600 }} title={name}>{name}</Typography>
      {!compact && <Typography component="span" noWrap variant="caption" color="text.secondary" sx={{ display: 'block' }} title={account.email}>E-mail : {account.email || 'Non renseigné'}</Typography>}
    </Box>
  </Box>
}
