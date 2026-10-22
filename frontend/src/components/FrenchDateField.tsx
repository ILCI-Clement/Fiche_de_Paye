import { useRef } from 'react'
import dayjs from 'dayjs'
import { DatePicker } from '@mui/x-date-pickers/DatePicker'
import { CalendarMonthHeader } from './CalendarMonthHeader'

/** French display and calendar; API values remain ISO dates. */
export function FrenchDateField({ label, value, onChange, disabled, required, fullWidth, min, error, helperText }: {
  label: string; value: string; onChange: (value: string) => void; disabled?: boolean;
  required?: boolean; fullWidth?: boolean; min?: string; error?: boolean; helperText?: string;
}) {
  const input = useRef<HTMLInputElement>(null)
  return <DatePicker label={label} value={value ? dayjs(value) : null} format="DD/MM/YYYY"
    views={['day']} slots={{ calendarHeader: CalendarMonthHeader }}
    disabled={disabled} minDate={min ? dayjs(min) : undefined} inputRef={input}
    onChange={(date, context) => {
      const invalid = Boolean(context.validationError)
      input.current?.setCustomValidity(invalid ? 'Veuillez saisir une date valide au format JJ/MM/AAAA.' : '')
      // Keep invalid input invalid rather than silently submitting the previous date.
      onChange(date ? (invalid ? 'Invalid Date' : date.format('YYYY-MM-DD')) : '')
    }}
    onError={reason => input.current?.setCustomValidity(reason ? 'Veuillez saisir une date valide au format JJ/MM/AAAA.' : '')}
    slotProps={{ textField: { required, fullWidth, ...(error ? { error } : {}), helperText } }} />
}
