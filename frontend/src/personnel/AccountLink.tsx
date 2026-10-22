import { useState } from 'react'
import { Alert, Box, Button, CircularProgress, MenuItem, TextField, Typography } from '@mui/material'
import { useQuery } from '@tanstack/react-query'
import { personnelApi } from './api'
import { useRegistryViewer } from './context'
import { personDisplayName } from '../personDisplayName'
import type { RegistryAccount } from './api'

export function AccountLink({ token, employeeId, value, currentDisplayName, busy, onChange }: { token: string; employeeId?: number; value: string | null; currentDisplayName?: string | null; busy: boolean; onChange: (value: string | null, displayName: string) => void }) {
  const [search, setSearch] = useState('')
  const [query, setQuery] = useState('')
  const [chosen, setChosen] = useState<RegistryAccount | null>(null)
  const viewer = useRegistryViewer()
  const accounts = useQuery({ queryKey: ['personnel-register', viewer, 'accounts', query], queryFn: ({ signal }) => personnelApi.accounts(token, query, signal), retry: false })
  const options = accounts.data?.accounts || []
  const fallback = value && !options.some(account => account.username === value)
  return <Box sx={{ gridColumn: '1 / -1', borderTop: 1, borderColor: 'divider', pt: 2 }}>
    <Typography variant="subtitle2" sx={{ mb: 1.5 }}>Compte de connexion · facultatif</Typography>
    <Box sx={{ display: 'flex', gap: 1, flexWrap: 'wrap', mb: 1.5 }}>
      <TextField label="Rechercher un compte par nom ou e-mail" value={search} disabled={busy} onChange={event => setSearch(event.target.value)} slotProps={{ htmlInput: { maxLength: 200 } }} sx={{ flex: 1, minWidth: 200 }} onKeyDown={event => { if (event.key === 'Enter') { event.preventDefault(); setQuery(search.trim()) } }} />
      <Button disabled={busy || accounts.isFetching} onClick={() => setQuery(search.trim())}>Rechercher</Button>
      {accounts.isFetching && <CircularProgress size={20} sx={{ alignSelf: 'center' }} />}
    </Box>
    {accounts.error && <Alert severity="error" sx={{ mb: 1 }}>{accounts.error.message}<Button onClick={() => void accounts.refetch()}>Réessayer</Button></Alert>}
    <TextField select label="Compte associé" fullWidth value={value || ''} disabled={busy || accounts.isPending || !!accounts.error} onChange={event => { const account = options.find(person => person.username === event.target.value) || null; setChosen(account); onChange(event.target.value || null, account ? personDisplayName(account) : '') }} helperText="Un compte ne peut être associé qu’à une seule personne. L’association ne modifie ni son mot de passe ni ses droits.">
      <MenuItem value="">Aucun compte associé</MenuItem>
      {fallback && <MenuItem value={value}>{chosen?.username === value ? personDisplayName(chosen) : personDisplayName({ display_name: currentDisplayName })} · association actuelle</MenuItem>}
      {options.map(account => {
        const unavailable = account.linked_employee_id !== null && account.linked_employee_id !== employeeId
        return <MenuItem key={account.username} value={account.username} disabled={unavailable}>{personDisplayName(account)} · {account.email}{unavailable ? ' · déjà associé' : account.is_active === false ? ' · connexion désactivée' : ''}</MenuItem>
      })}
    </TextField>
    {accounts.data?.has_more && <Typography variant="caption" color="text.secondary">Seuls les {accounts.data.limit} premiers résultats sont affichés. Affinez la recherche.</Typography>}
  </Box>
}
