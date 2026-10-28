import { Chip } from '@mui/material'
import { alpha } from '@mui/material/styles'

// Stable account-independent colours: renaming a category keeps its appearance.
const tones = [
  ['#1e64b8', '#8cbcff'], ['#076e68', '#72d6ca'], ['#7652ac', '#c7a7f3'],
  ['#986000', '#efc373'], ['#a43e68', '#efa3c1'], ['#426b25', '#b0d785'],
] as const

function categoryTone(code: string) {
  if (code === 'transport') return 0
  if (code === 'autre') return 2
  let hash = 0
  for (const char of code) hash = (hash * 31 + char.charCodeAt(0)) >>> 0
  return [1, 3, 4, 5][hash % 4]
}

export function DocumentTypeBadge({ code, label, inactive = false }: { code?: string | null; label: string; inactive?: boolean }) {
  return <Chip size="small" variant="outlined" label={label} title={label} sx={theme => {
    const dark = theme.palette.mode === 'dark'
    const muted = inactive || !code || code === 'all' || code === 'unset'
    const color = muted ? theme.palette.text.secondary : tones[categoryTone(code)][dark ? 1 : 0]
    return { color, bgcolor: alpha(color, dark ? .16 : .1), borderColor: alpha(color, dark ? .5 : .38),
      borderWidth: 1, borderStyle: inactive ? 'dashed' : 'solid', borderRadius: '8px', height: 25,
      maxWidth: '100%', minWidth: 0, fontSize: 12, fontWeight: 500,
      '& .MuiChip-label': { px: 1.1, overflow: 'hidden', textOverflow: 'ellipsis' } }
  }} />
}
