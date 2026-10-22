import { Button, MenuItem, TextField, ToggleButton, ToggleButtonGroup } from '@mui/material'
import { WorkspaceToolbar } from '../components/WorkspaceToolbar'
import { AccountOption } from './AccountOption'
import { accountOptionLabel } from './types'
import type { Account } from './types'
import { PeriodToolbar } from './PeriodToolbar'

export function AttendanceToolbar({ month, year, busy, account, accounts, onPeriod, onAccount, onOpen, archived, onArchiveFilter }: {
  month: number; year: number; busy: boolean; account: string; accounts: Account[];
  onPeriod: (month: number, year: number) => void; onAccount: (value: string) => void; onOpen: () => void;
  archived: boolean; onArchiveFilter: (archived: boolean) => void;
}) {
  return <WorkspaceToolbar>
    <PeriodToolbar label="Mois des fiches de présence" month={month} year={year} disabled={busy} onChange={onPeriod} />
    <ToggleButtonGroup size="small" exclusive value={archived ? 'archived' : 'active'} disabled={busy} aria-label="Dossier des fiches" onChange={(_, value) => { if (value) onArchiveFilter(value === 'archived') }}><ToggleButton value="active">Normal</ToggleButton><ToggleButton value="archived">Archives</ToggleButton></ToggleButtonGroup>
    <TextField select label="Sélectionner une personne" value={account} disabled={busy} onChange={event => onAccount(event.target.value)} sx={{ flex: '1 1 260px', minWidth: { xs: 0, sm: 240 }, maxWidth: { xs: 'none', lg: 540 }, '& .MuiInputBase-root': { minHeight: 44 }, '& .MuiInputLabel-root': { fontSize: 13 } }} slotProps={{ select: {
      renderValue: value => { const selected = accounts.find(item => item.username === value); return selected ? <AccountOption account={selected} compact /> : '' },
      MenuProps: { slotProps: { paper: { sx: { maxHeight: 420, maxWidth: 'calc(100vw - 24px)' } } } },
    } }}>
      <MenuItem value="">Sélectionner une personne</MenuItem>{accounts.map(item => <MenuItem key={item.username} value={item.username} aria-label={accountOptionLabel(item)} sx={{ py: 1, minHeight: 60 }}><AccountOption account={item} /></MenuItem>)}
    </TextField>
    <Button variant="contained" disabled={!account || busy} onClick={onOpen} sx={{ flexShrink: 0, whiteSpace: 'nowrap' }}>Ouvrir / créer sa fiche</Button>
  </WorkspaceToolbar>
}
