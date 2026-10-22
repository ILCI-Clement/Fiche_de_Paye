import { Box, Chip, Typography } from '@mui/material'
import { alpha } from '@mui/material/styles'
import { personDisplayName } from '../personDisplayName'
import { accessRoleLabel, displayFunctionRoles } from '../rolePolicy'
import { personnelTypes } from './api'
import type { AdminUser } from './api'
import { accountOptionColumns as columns } from './accountOptionLayout'

function AccessBadge({ roles }: { roles: readonly string[] }) {
  const label = accessRoleLabel(roles)
  return <Chip size="small" label={label} sx={theme => ({ maxWidth: '100%', height: 25, borderRadius: 1, bgcolor: label === 'Standard' ? 'action.hover' : alpha(theme.palette.primary.main, .12), color: label === 'Standard' ? 'text.secondary' : 'primary.main', '& .MuiChip-label': { px: 1, fontSize: 11, fontWeight: 500 } })} />
}

/** The same grid widths are used for the heading and every account option. */
export function AccountOptionHeader() {
  return <Box aria-hidden="true" sx={{ display: 'grid', gridTemplateColumns: columns, columnGap: { xs: 1, sm: 2 }, width: '100%', minWidth: 0, color: 'text.secondary', fontSize: 11, fontWeight: 500, lineHeight: 1.4 }}>
    <span>Identité</span><span>Fonction / personnel</span><span>Accès</span>
  </Box>
}

export function AccountOption({ person, compact = false, detail = '' }: { person: AdminUser; compact?: boolean; detail?: string }) {
  const name = personDisplayName(person)
  const access = accessRoleLabel(person.role_tags)
  if (compact) return <Box sx={{ display: 'flex', alignItems: 'center', gap: 1, width: '100%', minWidth: 0 }}>
    <Typography noWrap sx={{ fontSize: 13, flex: 1, minWidth: 0 }}>{name}</Typography>
    {access !== 'Standard' && <Box sx={{ flexShrink: 0 }}><AccessBadge roles={person.role_tags} /></Box>}
  </Box>
  const personnel = personnelTypes.find(([id]) => id === person.employee_type)?.[1] || 'Non renseigné'
  const functions = displayFunctionRoles(person.role_tags)
  const job = person.job_title?.trim()
  const functionLabel = job || functions || personnel
  const secondaryFunction = job ? [functions, personnel].filter(Boolean).join(' · ') : functions ? personnel : ''
  const status = person.login_deleted_at ? 'Connexion supprimée' : person.is_active === false ? 'Désactivé' : ''
  return <Box sx={{ display: 'grid', gridTemplateColumns: columns, columnGap: { xs: 1, sm: 2 }, alignItems: 'center', width: '100%', minWidth: 0 }}>
    <Box sx={{ minWidth: 0 }}>
      <Typography sx={{ fontSize: 13, fontWeight: 500, whiteSpace: 'normal', overflowWrap: 'anywhere', lineHeight: 1.4 }}>{name}</Typography>
      <Typography sx={{ fontSize: 11, color: 'text.secondary', whiteSpace: 'normal', overflowWrap: 'anywhere', lineHeight: 1.4 }}>{person.email || 'E-mail non renseigné'}</Typography>
      {detail && <Typography sx={{ fontSize: 11, color: 'text.secondary', whiteSpace: 'normal', overflowWrap: 'anywhere', lineHeight: 1.4 }}>{detail}</Typography>}
      {status && <Typography sx={{ fontSize: 11, color: 'text.secondary', whiteSpace: 'normal', lineHeight: 1.4 }}>{status}</Typography>}
    </Box>
    <Box sx={{ minWidth: 0 }}>
      <Typography sx={{ fontSize: 12, whiteSpace: 'normal', overflowWrap: 'anywhere', lineHeight: 1.4 }}>{functionLabel}</Typography>
      {secondaryFunction && <Typography sx={{ fontSize: 11, color: 'text.secondary', whiteSpace: 'normal', overflowWrap: 'anywhere', lineHeight: 1.4 }}>{secondaryFunction}</Typography>}
    </Box>
    <Box sx={{ minWidth: 0, justifySelf: 'start' }}><AccessBadge roles={person.role_tags} /></Box>
  </Box>
}
