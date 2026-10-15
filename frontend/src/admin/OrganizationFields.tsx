import { MenuItem, TextField } from '@mui/material'
import { personnelTypes } from './api'
import type { AdminUser, OrganizationInput } from './api'
import type { Group } from '../team/api'
import { DepartmentBadge, DepartmentBadges } from '../components/DepartmentBadge'

export function OrganizationFields({ value, onChange, users, groups, username = '', disabled }: {
  value: OrganizationInput; onChange: (value: OrganizationInput) => void; users: AdminUser[]; groups: Group[]; username?: string; disabled: boolean
}) {
  const managers = users.filter(p => p.username !== username && p.role_tags.some(tag => ['Admin', 'Responsable'].includes(tag)))
  const allowed = groups.filter(g => g.is_active || value.group_ids.includes(g.id) || value.managed_group_ids.includes(g.id))
  const multi = (input: string | string[]) => typeof input === 'string' ? input.split(',') : input
  return <>
    <TextField select label="Étiquettes d’identité" value={value.role_tags} disabled={disabled} slotProps={{ select: { multiple: true } }} onChange={e => onChange({ ...value, role_tags: multi(e.target.value) })}>{['Employe', 'Responsable', 'Admin'].map(role => <MenuItem key={role} value={role}>{role}</MenuItem>)}</TextField>
    <TextField select label="Type de personnel" value={value.employee_type} disabled={disabled} onChange={e => onChange({ ...value, employee_type: e.target.value })}>{personnelTypes.map(([id, label]) => <MenuItem key={id} value={id}>{label}</MenuItem>)}</TextField>
    {(['group_ids', 'managed_group_ids'] as const).map(field => <TextField key={field} select label={field === 'group_ids' ? 'Départements de travail' : 'Départements gérés'} value={value[field]} disabled={disabled} slotProps={{ select: { multiple: true, renderValue: () => <DepartmentBadges names={value[field].map(id => groups.find(g => g.id === id)?.name || 'Unknown')} /> } }} onChange={e => onChange({ ...value, [field]: multi(e.target.value).map(Number) })}>{allowed.map(g => <MenuItem key={g.id} value={g.id} sx={{ gap: 1 }} aria-label={g.name}><DepartmentBadge name={g.name} />{!g.is_active && ' · inactif'}</MenuItem>)}</TextField>)}
    <TextField select label="Responsable direct" value={value.manager_id || ''} disabled={disabled} onChange={e => onChange({ ...value, manager_id: e.target.value || null })}><MenuItem value="">Aucun responsable</MenuItem>{managers.map(p => <MenuItem key={p.username} value={p.username}>{p.username}</MenuItem>)}</TextField>
  </>
}
