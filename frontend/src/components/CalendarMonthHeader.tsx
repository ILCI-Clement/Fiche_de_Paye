import { Box, IconButton, TextField } from '@mui/material'
import ChevronLeftRounded from '@mui/icons-material/ChevronLeftRounded'
import ChevronRightRounded from '@mui/icons-material/ChevronRightRounded'
import dayjs from 'dayjs'
import type { PickersCalendarHeaderProps } from '@mui/x-date-pickers/PickersCalendarHeader'

/** Month and year use ordinary vertical menus, including on mobile pickers. */
export function CalendarMonthHeader({ currentMonth, onMonthChange, minDate, maxDate, disablePast, disableFuture, disabled, labelId }: PickersCalendarHeaderProps) {
  const current = dayjs(currentMonth)
  const today = dayjs()
  const lower = disablePast && today.isAfter(minDate) ? today : dayjs(minDate)
  const upper = disableFuture && today.isBefore(maxDate) ? today : dayjs(maxDate)
  const first = lower.startOf('month')
  const last = upper.startOf('month')
  function change(value: dayjs.Dayjs) {
    const month = value.startOf('month')
    onMonthChange(month.isBefore(first) ? first : month.isAfter(last) ? last : month)
  }
  // Native lists do not create another modal/portal outside the picker.
  // That prevents its focus trap and click-away handler from closing the calendar.
  return <Box sx={{ display: 'flex', alignItems: 'center', gap: .5, px: 1, py: 1.5 }}>
    <Box component="span" id={labelId} sx={{ position: 'absolute', width: 1, height: 1, overflow: 'hidden', clipPath: 'inset(50%)' }}>{current.format('MMMM YYYY')}</Box>
    <IconButton size="small" aria-label="Mois précédent" disabled={disabled || !current.startOf('month').isAfter(first)} onClick={() => change(current.subtract(1, 'month'))}><ChevronLeftRounded fontSize="small" /></IconButton>
    <TextField select size="small" label="Mois" value={current.month()} disabled={disabled} onChange={event => change(current.month(Number(event.target.value)))} sx={{ flex: 1, minWidth: 0 }} slotProps={{ select: { native: true }, inputLabel: { shrink: true }, htmlInput: { 'aria-label': 'Mois du calendrier' } }}>
      {Array.from({ length: 12 }, (_, month) => {
        const value = current.month(month).startOf('month')
        return <option key={month} value={month} disabled={value.isBefore(first) || value.isAfter(last)}>{value.format('MMMM')}</option>
      })}
    </TextField>
    <TextField select size="small" label="Année" value={current.year()} disabled={disabled} onChange={event => change(current.year(Number(event.target.value)))} sx={{ width: 92 }} slotProps={{ select: { native: true }, inputLabel: { shrink: true }, htmlInput: { 'aria-label': 'Année du calendrier' } }}>
      {Array.from({ length: Math.max(0, upper.year() - lower.year() + 1) }, (_, index) => lower.year() + index).map(year => <option key={year} value={year}>{year}</option>)}
    </TextField>
    <IconButton size="small" aria-label="Mois suivant" disabled={disabled || !current.startOf('month').isBefore(last)} onClick={() => change(current.add(1, 'month'))}><ChevronRightRounded fontSize="small" /></IconButton>
  </Box>
}
