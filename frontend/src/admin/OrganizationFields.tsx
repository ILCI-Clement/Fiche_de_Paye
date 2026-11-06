import { ListSubheader, MenuItem, TextField, Typography } from '@mui/material'
import { AccountOption, AccountOptionHeader } from './AccountOption'
import { accountMenuProps } from './accountOptionLayout'
import { personnelTypes } from './api'
import type { AdminUser, OrganizationInput } from './api'
import type { Group } from '../team/api'
import { DepartmentBadge, DepartmentBadges } from '../components/DepartmentBadge'
import { canManageTeam, roleLabel } from '../rolePolicy'
import { businessFunction, withBusinessFunction } from './accountRoles'
import { SystemAccessFields } from './SystemAccessFields'
import { accountStatus } from './accountStatus'

export function OrganizationFields({ value, onChange, users, groups, username = '', disabled, preserveAdmin = false, allowSuperAdmin = false, preserveSuperAdmin = false, hidePersonnelType = false, hideSystemAccess = false }: {
  value: OrganizationInput; onChange: (value: OrganizationInput) => void; users: AdminUser[]; groups: Group[]; username?: string; disabled: boolean; preserveAdmin?: boolean; allowSuperAdmin?: boolean; preserveSuperAdmin?: boolean; hidePersonnelType?: boolean; hideSystemAccess?: boolean
}) {
  const managers = users.filter(p => p.username !== username && (accountStatus(p) === 'active' || p.username === value.manager_id) && canManageTeam(p))
  const allowed = groups.filter(g => g.is_active || value.group_ids.includes(g.id) || value.managed_group_ids.includes(g.id))
  const multi = (input: string | string[]) => typeof input === 'string' ? input.split(',') : input
  if (value.role_tags.includes('SuperAdmin')) return <>
    <Typography variant="body2" color="text.secondary" sx={{ gridColumn: '1 / -1' }}>Compte système : aucune fiche du personnel, aucun contrat, département ou responsable requis.</Typography>
    {!hideSystemAccess && <SystemAccessFields roles={value.role_tags} onChange={roles => onChange({ ...value, role_tags: roles })}
      disabled={disabled} allowSuperAdmin={allowSuperAdmin} preserveAdmin={preserveAdmin} preserveSuperAdmin={preserveSuperAdmin} />}
  </>
  return <>
    <TextField select label="Fonction" value={businessFunction(value.role_tags)} disabled={disabled} onChange={e => {
      const roles = withBusinessFunction(value.role_tags, e.target.value)
      onChange({ ...value, role_tags: roles, managed_group_ids: roles.includes('Responsable') ? value.managed_group_ids : [] })
    }}>
      <MenuItem value="">Sans fonction métier</MenuItem>
      {['Employe', 'Responsable'].map(role => <MenuItem key={role} value={role}>{roleLabel(role)}</MenuItem>)}
    </TextField>
    {!hideSystemAccess && <SystemAccessFields roles={value.role_tags} onChange={roles => onChange({ ...value, role_tags: roles })}
      disabled={disabled} allowSuperAdmin={allowSuperAdmin} preserveAdmin={preserveAdmin} preserveSuperAdmin={preserveSuperAdmin} />}
    {!hidePersonnelType && <TextField select label="Type de personnel" value={value.employee_type} disabled={disabled} onChange={e => onChange({ ...value, employee_type: e.target.value })}>{personnelTypes.map(([id, label]) => <MenuItem key={id} value={id}>{label}</MenuItem>)}</TextField>}
    {(['group_ids', 'managed_group_ids'] as const).map(field => <TextField key={field} select label={field === 'group_ids' ? 'Départements de travail' : 'Départements gérés'} value={value[field]} disabled={disabled || (field === 'managed_group_ids' && !value.role_tags.includes('Responsable'))} title={field === 'managed_group_ids' && !value.role_tags.includes('Responsable') ? 'Réservé à la fonction Responsable.' : undefined} slotProps={{ select: { multiple: true, renderValue: () => <DepartmentBadges names={value[field].map(id => groups.find(g => g.id === id)?.name || 'Unknown')} /> } }} onChange={e => onChange({ ...value, [field]: multi(e.target.value).map(Number) })}>{allowed.map(g => <MenuItem key={g.id} value={g.id} sx={{ gap: 1 }} aria-label={g.name}><DepartmentBadge name={g.name} />{!g.is_active && ' · inactif'}</MenuItem>)}</TextField>)}
    <TextField data-completion-field="manager" select label="Responsable direct" value={value.manager_id || ''} disabled={disabled} slotProps={{ select: { MenuProps: accountMenuProps, renderValue: selected => { const person = managers.find(p => p.username === selected); return person ? <AccountOption person={person} compact /> : 'Aucun responsable' } } }} onChange={e => onChange({ ...value, manager_id: e.target.value || null })}><MenuItem value="">Aucun responsable</MenuItem><ListSubheader><AccountOptionHeader /></ListSubheader>{managers.map(p => <MenuItem key={p.username} value={p.username} disabled={p.is_active === false}><AccountOption person={p} /></MenuItem>)}</TextField>
  </>
}
