import { Box } from '@mui/material'
import { useRef, useState } from 'react'
import dayjs from 'dayjs'
import { DatePicker } from '@mui/x-date-pickers/DatePicker'
import { CalendarMonthHeader } from './CalendarMonthHeader'

/** French display and calendar; API values remain ISO dates. */
export function FrenchDateField({ label, value, onChange, disabled, required, fullWidth, min, error, helperText, completionField }: { completionField?: string;
  label: string; value: string; onChange: (value: string) => void; disabled?: boolean;
  required?: boolean; fullWidth?: boolean; min?: string; error?: boolean; helperText?: string;
}) {
  const input = useRef<HTMLInputElement>(null)
  const [draft, setDraft] = useState(() => ({ published: value, date: value ? dayjs(value) : null }))
  // Keep the picker's exact draft while its serialized value is echoed by the form.
  // Re-parsing "Invalid Date" during year entry clears the already-entered sections.
  if (value !== draft.published) {
    setDraft({ published: value, date: value ? dayjs(value) : null })
  }
  const picker = <DatePicker label={label} value={draft.date} format="DD/MM/YYYY"
    views={['day']} slots={{ calendarHeader: CalendarMonthHeader }}
    disabled={disabled} minDate={min ? dayjs(min) : undefined} inputRef={input}
    onChange={(date, context) => {
      const invalid = Boolean(context.validationError)
      input.current?.setCustomValidity(invalid ? 'Veuillez saisir une date valide au format JJ/MM/AAAA.' : '')
      // Keep invalid input invalid rather than silently submitting the previous date.
      const published = date ? (invalid ? 'Invalid Date' : date.format('YYYY-MM-DD')) : ''
      setDraft({ published, date })
      onChange(published)
    }}
    onError={reason => input.current?.setCustomValidity(reason ? 'Veuillez saisir une date valide au format JJ/MM/AAAA.' : '')}
    slotProps={{ textField: { required, fullWidth, ...(error ? { error } : {}), helperText } }} />
  return completionField ? <Box data-completion-field={completionField}>{picker}</Box> : picker
}
