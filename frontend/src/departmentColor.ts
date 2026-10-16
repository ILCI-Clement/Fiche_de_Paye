export function departmentColor(name: string) {
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

