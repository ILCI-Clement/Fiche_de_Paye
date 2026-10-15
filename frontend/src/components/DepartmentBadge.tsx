import { Box } from '@mui/material'
import { alpha } from '@mui/material/styles'

function departmentColor(name: string) {
  const normalized = name.normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().replace(/^groupe\s+/, '').trim()
  const known: Record<string, string> = {
    'direction generale': '#94a3b8', rh: '#f472b6', finance: '#fbbf24',
    marketing: '#a78bfa', pedagogie: '#60a5fa', 'r&d': '#22d3ee', scolarite: '#fb923c',
    unknown: '#94a3b8',
  }
  if (known[normalized]) return known[normalized]
  const colors = ['#60a5fa', '#a78bfa', '#fbbf24', '#22d3ee', '#f472b6', '#fb923c']
  const hash = [...normalized].reduce((value, char) => (value * 31 + char.charCodeAt(0)) >>> 0, 0)
  return colors[hash % colors.length]
}

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
