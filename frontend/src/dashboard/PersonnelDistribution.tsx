import { useState } from 'react'
import { Alert, Box, Button, ButtonBase, CircularProgress, Paper, Tooltip, Typography } from '@mui/material'
import { alpha, useTheme } from '@mui/material/styles'
import { useQuery } from '@tanstack/react-query'
import { teamApi } from '../team/api'
import { departmentColor } from '../departmentColor'

export function PersonnelDistribution({ token, username }: { token: string; username: string }) {
  const theme = useTheme()
  const [active, setActive] = useState<string | null>(null)
  const query = useQuery({ queryKey: ['organization', username], queryFn: ({ signal }) => teamApi.directory(token, signal), retry: false })
  const groups = query.data?.groups.filter(group => group.is_active) || []
  const buckets = groups.map(group => ({ key: String(group.id), name: group.name, count: 0, color: departmentColor(group.name) }))
  const unassigned = { key: 'unassigned', name: 'Sans département actif', count: 0, color: '#94a3b8' }
  const multiple = { key: 'multiple', name: 'Affectations multiples à vérifier', count: 0, color: '#c49b56' }
  const people = [...new Map((query.data?.users || []).map(person => [person.username, person])).values()]
  for (const person of people) {
    const ids = [...new Set(person.group_ids)].filter(id => groups.some(group => group.id === id))
    if (!ids.length) unassigned.count++
    else if (ids.length > 1) multiple.count++
    else buckets.find(bucket => bucket.key === String(ids[0]))!.count++
  }
  const slices = [...buckets, unassigned, multiple].filter(bucket => bucket.count > 0)
  const total = people.length
  const detail = slices.find(slice => slice.key === active)
  const percent = (count: number) => new Intl.NumberFormat('fr-FR', { style: 'percent', maximumFractionDigits: 1 }).format(total ? count / total : 0)
  let offset = 0
  return <Paper component="section" aria-labelledby="personnel-distribution-title" sx={{ p: { xs: 2, sm: 3 }, mb: 3, borderRadius: 3 }}>
    <Typography id="personnel-distribution-title" variant="h6" sx={{ fontWeight: 600 }}>Répartition des effectifs</Typography>
    <Typography variant="body2" color="text.secondary" sx={{ mt: .5 }}>Comptes du répertoire · un compte compté une seule fois, selon son département d’appartenance.</Typography>
    {query.isPending ? <CircularProgress size={28} sx={{ m: 3 }} /> : query.error ? <Alert severity="error" sx={{ mt: 2 }}>{query.error.message}<Button onClick={() => void query.refetch()}>Réessayer</Button></Alert>
      : !total ? <Typography color="text.secondary" sx={{ mt: 2 }}>Aucune personne dans le répertoire.</Typography> : <>
        <Box sx={{ display: 'grid', gridTemplateColumns: { xs: '1fr', md: 'minmax(220px, 300px) minmax(0, 1fr)' }, alignItems: 'center', gap: 3, mt: 2 }}>
          <Box sx={{ width: '100%', maxWidth: 280, mx: 'auto', '& svg:focus, & svg circle:focus': { outline: 'none' } }}>
            <svg viewBox="0 0 240 240" style={{ width: '100%', display: 'block', overflow: 'visible' }} role="group" aria-label="Répartition des comptes par département">
              <circle cx="120" cy="120" r="90" fill="none" stroke={theme.palette.divider} strokeWidth="32" />
              {slices.map(slice => {
                const size = slice.count / total * 100
                const start = offset; offset += size
                const label = `${slice.name} : ${slice.count} personne(s), ${percent(slice.count)}`
                return <Tooltip key={slice.key} title={label} arrow><circle cx="120" cy="120" r="90" fill="none" pathLength="100"
                  stroke={slice.color} strokeWidth={active === slice.key ? 38 : 32} strokeDasharray={`${size} ${100 - size}`} strokeDashoffset={-start}
                  transform="rotate(-90 120 120)" tabIndex={0} role="button" aria-label={label}
                  onMouseEnter={() => setActive(slice.key)} onMouseLeave={() => setActive(null)} onFocus={() => setActive(slice.key)} onBlur={() => setActive(null)}
                  onClick={() => setActive(slice.key)} onKeyDown={event => { if (event.key === 'Enter' || event.key === ' ') { event.preventDefault(); setActive(slice.key) } }}
                  style={{ cursor: 'pointer', opacity: active && active !== slice.key ? .45 : 1, transition: 'opacity .15s, stroke-width .15s' }} /></Tooltip>
              })}
              <text x="120" y="120" textAnchor="middle" fill={theme.palette.text.primary} style={{ fontSize: 36, fontWeight: 600 }}>{total}</text>
              <text x="120" y="145" textAnchor="middle" fill={theme.palette.text.secondary} style={{ fontSize: 13 }}>personnes</text>
            </svg>
          </Box>
          <Box component="ul" aria-label="Effectifs par département" sx={{ listStyle: 'none', p: 0, m: 0 }}>
            {slices.map(slice => <Box component="li" key={slice.key}><ButtonBase onClick={() => setActive(slice.key)} onMouseEnter={() => setActive(slice.key)} onMouseLeave={() => setActive(null)} onFocus={() => setActive(slice.key)} onBlur={() => setActive(null)}
              sx={{ width: '100%', display: 'flex', gap: 1.5, alignItems: 'center', textAlign: 'left', px: 1.5, py: 1.25, borderRadius: 2, bgcolor: active === slice.key ? alpha(slice.color, .1) : 'transparent', '&:hover': { bgcolor: alpha(slice.color, .1) } }}>
              <Box component="span" sx={{ width: 10, height: 10, borderRadius: '50%', bgcolor: slice.color, flexShrink: 0 }} />
              <Typography variant="body2" sx={{ flex: 1, overflowWrap: 'anywhere' }}>{slice.name}</Typography>
              <Typography variant="body2" sx={{ fontWeight: 500, fontVariantNumeric: 'tabular-nums' }}>{slice.count}</Typography>
              <Typography variant="caption" color="text.secondary" sx={{ minWidth: 50, textAlign: 'right' }}>{percent(slice.count)}</Typography>
            </ButtonBase></Box>)}
          </Box>
        </Box>
        <Typography variant="body2" color="text.secondary" aria-live="polite" sx={{ mt: 1, minHeight: 44 }}>{detail ? `${detail.name} · ${detail.count} personne(s) · ${percent(detail.count)}` : 'Survolez une portion ou sélectionnez un département pour afficher le détail.'}</Typography>
        {!!multiple.count && <Alert severity="warning">Les comptes affectés à plusieurs départements sont regroupés séparément pour éviter les doubles comptes.</Alert>}
      </>}
  </Paper>
}
