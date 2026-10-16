import { Box } from '@mui/material'
import { alpha } from '@mui/material/styles'

import { departmentColor } from '../departmentColor'

/** Fixed-width department column, shared by selectors, directories and cards. */
export function DepartmentBadge({ name = 'Unknown' }: { name?: string }) {
  const fullName = name.trim() || 'Unknown'
  const label = fullName.replace(/^groupe\s+/i, '')
  const color = departmentColor(fullName)
  return <Box component="span" title={fullName} className="department-badge" sx={theme => ({
    backgroundColor: alpha(color, theme.palette.mode === 'dark' ? .14 : .16),
    color: theme.palette.mode === 'dark' ? color : '#31333f',
    borderColor: alpha(color, theme.palette.mode === 'dark' ? .35 : .65),
    '&::before': { backgroundColor: color },
  })}><span>{label}</span></Box>
}

export function DepartmentBadges({ names }: { names: string[] }) {
  const departments = [...new Set(names.map(name => name.trim()).filter(Boolean))]
  return <Box component="span" className="department-badges">{(departments.length ? departments : ['Unknown']).map(name => <DepartmentBadge key={name} name={name} />)}</Box>
}
