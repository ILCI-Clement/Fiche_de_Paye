import { useEffect, useState } from 'react'
import { Alert, Box, Button, CircularProgress, MenuItem, TextField, Typography } from '@mui/material'
import { useQuery } from '@tanstack/react-query'
import { attendanceApi } from './api'
import type { Cell } from './types'
import { PeriodSelector } from './PeriodSelector'
import { useCalendarSelection } from './useCalendarSelection'
import { useTheme } from '@mui/material/styles'
import { calendarStatus, calendarStatuses } from './calendarStatus'

const statuses = [...calendarStatuses, 'Réinitialiser']

export function AttendanceCalendar({ token, id, month, year, busy, version, active, onPeriod, onApply, onSelectionPending }: {
  token: string; id: string; month: number; year: number; busy: boolean; version: number; active: boolean;
  onPeriod: (month: number, year: number) => void; onApply: (cells: Cell[], status: string, reason: string) => Promise<boolean | undefined>
  onSelectionPending: (pending: boolean) => void
}) {
  const { palette: { mode } } = useTheme()
  const { selection, clear, cellProps } = useCalendarSelection(busy)
  const pending = selection.length > 0
  useEffect(() => { onSelectionPending(pending); return () => onSelectionPending(false) }, [pending, onSelectionPending])
  const [status, setStatus] = useState('Congé payé')
  const [reason, setReason] = useState('')
  const query = useQuery({ queryKey: ['attendance-calendar', id, month, year, version],
    queryFn: ({ signal }) => attendanceApi.calendar(token, id, month, year, signal), enabled: active })
  return <Box>
      <PeriodSelector month={month} year={year} disabled={busy} onChange={onPeriod} />
      <Typography variant="body2" color="text.secondary" sx={{ my: 2 }}>Cliquez ou glissez sur les demi-journées pour sélectionner. Glissez depuis une case sélectionnée pour désélectionner, puis choisissez un statut et appliquez.</Typography>
      <Box className="calendar-toolbar">
        <TextField select label="Statut" value={status} onChange={e => setStatus(e.target.value)} disabled={busy} sx={{ minWidth: 160 }}>
          {statuses.map(s => <MenuItem key={s} value={s}>
            {s !== 'Réinitialiser' && <Box component="span" sx={{ width: 10, height: 10, borderRadius: '50%', bgcolor: calendarStatus(s, mode).foreground, mr: 1 }} />}{s}
          </MenuItem>)}
        </TextField>
        {status === 'Autre' && <TextField label="Précision" value={reason} onChange={e => setReason(e.target.value)} disabled={busy} />}
        <Button variant="contained" disabled={!selection.length || busy} onClick={async () => { if (await onApply(selection, status, reason)) clear() }}>Appliquer ({selection.length})</Button>
        <Button disabled={!selection.length || busy} onClick={clear}>Effacer la sélection</Button>
      </Box>
      {pending && <Typography role="status" variant="body2" color="primary.main" sx={{ mb: 2 }}>{selection.length} demi-journée(s) sélectionnée(s) : cliquez sur « Appliquer » pour enregistrer ce statut.</Typography>}
      <Box aria-label="Légende des statuts" sx={{ display: 'flex', flexWrap: 'wrap', gap: 1, mb: 2 }}>
        {calendarStatuses.map(s => {
          const color = calendarStatus(s, mode)
          return <Box key={s} component="span" sx={{ bgcolor: color.background, color: color.foreground, px: 1, py: .5, borderRadius: 1, fontSize: 12, fontWeight: 650 }}>{s}</Box>
        })}
        <Typography variant="caption" color="text.secondary" sx={{ alignSelf: 'center' }}>Contour bleu : sélection</Typography>
      </Box>
      {query.isPending && <CircularProgress size={24} sx={{ m: 2 }} />}
      {query.error && <Alert severity="error">{query.error.message}</Alert>}
      {query.data && <Box className="attendance-calendar-grid">
        {['Lun', 'Mar', 'Mer', 'Jeu', 'Ven', 'Sam', 'Dim'].map(day => <Typography key={day} className="calendar-weekday">{day}</Typography>)}
        {Array.from({ length: query.data.days[0]?.weekday || 0 }, (_, i) => <Box key={`empty-${i}`} />)}
        {query.data.days.map(day => <Box key={day.date} className={`calendar-day ${day.disabled ? 'outside-contract' : ''}`}>
          <Typography className="calendar-number">{day.number}</Typography>
          {(['matin', 'aprem'] as const).map(period => {
            const color = calendarStatus(day[period][0], mode)
            return <Button key={period} className={`calendar-half ${selection.some(c => c.date === day.date && c.period === period) ? 'selected' : ''}`}
            style={{ '--status-background': color.background, '--status-foreground': color.foreground } as React.CSSProperties}
            disabled={busy || day.disabled} title={`${day.date} · ${period === 'matin' ? 'Matin' : 'Après-midi'} · ${day[period].join(' · ')}`}
            aria-pressed={selection.some(c => c.date === day.date && c.period === period)}
            aria-label={`${day.date} ${period === 'matin' ? 'matin' : 'après-midi'} ${day[period][0]}`}
            {...cellProps({ date: day.date, period })}>{period === 'matin' ? 'M' : 'A'} · {color.short}</Button>
          })}
        </Box>)}
      </Box>}
  </Box>
}
