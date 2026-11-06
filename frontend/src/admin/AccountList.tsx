import { Box, ButtonBase, Paper, Typography } from '@mui/material'
import ChevronRightRounded from '@mui/icons-material/ChevronRightRounded'
import { AccountOption, AccountOptionHeader } from './AccountOption'
import { personDisplayName } from '../personDisplayName'
import type { AdminUser } from './api'

export function AccountList({ accounts, busy, onOpen }: {
  accounts: AdminUser[]; busy: boolean; onOpen: (username: string) => void
}) {
  return <Paper variant="outlined" sx={{ overflow: 'hidden', mb: 2 }}>
    <Box sx={{ px: 2, py: 1.25, pr: 6, bgcolor: 'action.hover' }}><AccountOptionHeader /></Box>
    {!accounts.length && <Typography color="text.secondary" sx={{ p: 3 }}>Aucun compte pour ce filtre.</Typography>}
    {accounts.map(person => <ButtonBase key={person.username} disabled={busy}
      aria-label={`Ouvrir le compte de ${personDisplayName(person)}`}
      onClick={() => onOpen(person.username)} sx={{ width: '100%', display: 'flex', textAlign: 'left',
        justifyContent: 'space-between', gap: 1, px: 2, py: 1.5, minHeight: 64,
        borderTop: 1, borderColor: 'divider', '&:hover': { bgcolor: 'action.hover' },
        '&.Mui-focusVisible': { outline: '2px solid', outlineColor: 'primary.main', outlineOffset: -2 } }}>
      <AccountOption person={person} />
      <ChevronRightRounded sx={{ flexShrink: 0, color: 'text.secondary', fontSize: 20 }} />
    </ButtonBase>)}
  </Paper>
}
