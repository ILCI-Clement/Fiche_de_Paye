import { Box } from '@mui/material'
import { PeriodSelector } from './PeriodSelector'

/** The shared period selector also works outside a page's .workspace styles. */
export function PeriodToolbar({ label = 'Choisir la période', ...props }: Parameters<typeof PeriodSelector>[0] & { label?: string }) {
  return <Box role="group" aria-label={label} sx={{ flexShrink: 0, width: { xs: '100%', sm: 'auto' }, minWidth: 0, maxWidth: '100%',
    '& .period-selector': { display: 'grid', gridTemplateColumns: { xs: '44px minmax(0, 1fr) 76px 44px', sm: '44px 130px 80px 44px' }, gap: .5, width: '100%', alignItems: 'center' },
    '& .period-selector > .MuiTextField-root': { minWidth: 0, width: '100%' },
    '& .period-selector > .MuiIconButton-root': { width: 44, height: 44, minWidth: 0, p: .25 },
    '& .MuiInputBase-root': { minHeight: 44 },
    '& .MuiInputBase-input, & .MuiInputLabel-root': { fontSize: 13 },
  }}><PeriodSelector {...props} /></Box>
}
