import { Box, ListSubheader, MenuItem, TextField, Typography } from '@mui/material'
import { AccountOption, AccountOptionHeader } from './AccountOption'
import { accountMenuProps } from './accountOptionLayout'
import { personnelTypes } from './api'
import type { AdminUser, OrganizationInput } from './api'
import type { Group } from '../team/api'
import { DepartmentBadge, DepartmentBadges } from '../components/DepartmentBadge'
import { accessRoleLabel, canManageTeam, displayFunctionRoles, roleLabel } from '../rolePolicy'

export function OrganizationFields({ value, onChange, users, groups, username = '', disabled, preserveAdmin = false, allowSuperAdmin = false, preserveSuperAdmin = false }: {
  value: OrganizationInput; onChange: (value: OrganizationInput) => void; users: AdminUser[]; groups: Group[]; username?: string; disabled: boolean; preserveAdmin?: boolean; allowSuperAdmin?: boolean; preserveSuperAdmin?: boolean
}) {
  const managers = users.filter(p => p.username !== username && (p.is_active !== false || p.username === value.manager_id) && canManageTeam(p))
  const allowed = groups.filter(g => g.is_active || value.group_ids.includes(g.id) || value.managed_group_ids.includes(g.id))
  const multi = (input: string | string[]) => typeof input === 'string' ? input.split(',') : input
  return <>
    <TextField select label="Fonction et droits" value={value.role_tags.filter(role => role !== 'Admin' || !value.role_tags.includes('SuperAdmin'))} disabled={disabled} helperText={preserveSuperAdmin ? 'Votre rôle Super Admin est protégé.' : preserveAdmin ? 'Votre rôle Admin est protégé contre un retrait accidentel.' : undefined} slotProps={{ select: { multiple: true, renderValue: () => <Box sx={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 1, minWidth: 0 }}><Typography noWrap sx={{ fontSize: 'inherit', minWidth: 0 }}>{displayFunctionRoles(value.role_tags) || 'Sans fonction'}</Typography><Typography noWrap sx={{ fontSize: 12, color: 'text.secondary', flexShrink: 0 }}>{accessRoleLabel(value.role_tags)}</Typography></Box> } }} onChange={e => {
      const roles = multi(e.target.value)
      if ((!preserveAdmin || roles.includes('Admin') || roles.includes('SuperAdmin')) && (!preserveSuperAdmin || roles.includes('SuperAdmin')) && (allowSuperAdmin || !roles.includes('SuperAdmin'))) onChange({ ...value, role_tags: roles })
    }}>{['Employe', 'Responsable', 'Admin', ...(allowSuperAdmin || value.role_tags.includes('SuperAdmin') ? ['SuperAdmin'] : [])].filter(role => role !== 'Admin' || !value.role_tags.includes('SuperAdmin')).map(role => <MenuItem key={role} value={role} disabled={(preserveAdmin && role === 'Admin') || (role === 'SuperAdmin' && (!allowSuperAdmin || preserveSuperAdmin))}>{roleLabel(role)}</MenuItem>)}</TextField>
    <TextField select label="Type de personnel" value={value.employee_type} disabled={disabled} onChange={e => onChange({ ...value, employee_type: e.target.value })}>{personnelTypes.map(([id, label]) => <MenuItem key={id} value={id}>{label}</MenuItem>)}</TextField>
    {(['group_ids', 'managed_group_ids'] as const).map(field => <TextField key={field} select label={field === 'group_ids' ? 'Départements de travail' : 'Départements gérés'} value={value[field]} disabled={disabled} slotProps={{ select: { multiple: true, renderValue: () => <DepartmentBadges names={value[field].map(id => groups.find(g => g.id === id)?.name || 'Unknown')} /> } }} onChange={e => onChange({ ...value, [field]: multi(e.target.value).map(Number) })}>{allowed.map(g => <MenuItem key={g.id} value={g.id} sx={{ gap: 1 }} aria-label={g.name}><DepartmentBadge name={g.name} />{!g.is_active && ' · inactif'}</MenuItem>)}</TextField>)}
    <TextField select label="Responsable direct" value={value.manager_id || ''} disabled={disabled} slotProps={{ select: { MenuProps: accountMenuProps, renderValue: selected => { const person = managers.find(p => p.username === selected); return person ? <AccountOption person={person} compact /> : 'Aucun responsable' } } }} onChange={e => onChange({ ...value, manager_id: e.target.value || null })}><MenuItem value="">Aucun responsable</MenuItem><ListSubheader><AccountOptionHeader /></ListSubheader>{managers.map(p => <MenuItem key={p.username} value={p.username} disabled={p.is_active === false}><AccountOption person={p} /></MenuItem>)}</TextField>
  </>
}
