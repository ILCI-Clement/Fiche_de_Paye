import { Box, IconButton, MenuItem, TextField } from '@mui/material'
import ChevronLeftRounded from '@mui/icons-material/ChevronLeftRounded'
import ChevronRightRounded from '@mui/icons-material/ChevronRightRounded'

const months = Array.from({ length: 12 }, (_, i) => new Intl.DateTimeFormat('fr-FR', { month: 'long' }).format(new Date(2026, i, 1)))

export function PeriodSelector({ month, year, onChange, disabled = false }: {
  month: number; year: number; onChange: (month: number, year: number) => void; disabled?: boolean
}) {
  function shift(delta: number) { const d = new Date(year, month - 1 + delta, 1); if (d.getFullYear() >= 2000 && d.getFullYear() <= 2100) onChange(d.getMonth() + 1, d.getFullYear()) }
  return <Box className="period-selector" sx={{ display: 'flex', gap: 1, alignItems: 'center', flexWrap: 'wrap' }}>
    <IconButton disabled={disabled} aria-label="Mois précédent" onClick={() => shift(-1)}><ChevronLeftRounded /></IconButton>
    <TextField select label="Mois" value={month} disabled={disabled} onChange={e => onChange(Number(e.target.value), year)} sx={{ minWidth: 150 }}>
      {months.map((label, index) => <MenuItem key={label} value={index + 1}>{label[0].toUpperCase() + label.slice(1)}</MenuItem>)}
    </TextField>
    <TextField label="Année" type="number" value={year} disabled={disabled} sx={{ width: 105 }} slotProps={{ htmlInput: { min: 2000, max: 2100 } }}
      onChange={e => { const value = Number(e.target.value); if (value >= 2000 && value <= 2100) onChange(month, value) }} />
    <IconButton disabled={disabled} aria-label="Mois suivant" onClick={() => shift(1)}><ChevronRightRounded /></IconButton>
  </Box>
}
