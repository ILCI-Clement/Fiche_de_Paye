import { useDraftGuard, useWorkspaceGuard } from '../draftGuard'
import { useState } from 'react'
import { Alert, Box, Button, Checkbox, CircularProgress, FormControlLabel, MenuItem, TextField, Typography } from '@mui/material'
import { useQuery } from '@tanstack/react-query'
import { adminApi, personnelTypes } from './api'
import type { Contract, RunAction } from './api'

export function ContractEditor({ token, username, busy, run }: { token: string; username: string; busy: boolean; run: RunAction }) {
  const contracts = useQuery({ queryKey: ['admin-contracts', username], queryFn: ({ signal }) => adminApi.contracts(token, username, signal), retry: false })
  const guard = useWorkspaceGuard()
  const [selected, setSelected] = useState<number | ''>('')
  if (contracts.isPending) return <CircularProgress size={24} />
  if (contracts.error) return <Alert severity="error">{contracts.error.message}<Button onClick={() => void contracts.refetch()}>Réessayer</Button></Alert>
  const current = contracts.data.contracts.find(c => c.id === selected)
  return <>
    <TextField select fullWidth label="Contrat à gérer" value={selected} disabled={busy} onChange={e => { const value = e.target.value === '' ? '' : Number(e.target.value); guard?.request(() => setSelected(value)) }}>
      <MenuItem value="">Nouveau contrat</MenuItem>{contracts.data.contracts.map(c => <MenuItem key={c.id} value={c.id}>{c.employee_type} · {c.start_date} → {c.end_date || 'CDI'}</MenuItem>)}
    </TextField>
    <Typography variant="body2" color="text.secondary" sx={{ my: 2 }}>Les contrats historiques sont conservés. Les périodes ne doivent pas se chevaucher.</Typography>
    <ContractForm key={`${selected}-${JSON.stringify(current)}`} current={current} busy={busy} save={data => run(() => adminApi.saveContract(token, username, current?.id ?? null, data), 'Contrat enregistré.')} />
  </>
}

function ContractForm({ current, busy, save }: { current?: Contract; busy: boolean; save: (data: Omit<Contract, 'id'>) => Promise<boolean> }) {
  const [kind, setKind] = useState(current?.employee_type || 'salarie')
  const [start, setStart] = useState(current?.start_date || '')
  const [end, setEnd] = useState(current?.end_date || '')
  const [cdi, setCdi] = useState(Boolean(current?.is_cdi))
  const [baseline, setBaseline] = useState(() => JSON.stringify([kind, start, end, cdi]))
  useDraftGuard(JSON.stringify([kind, start, end, cdi]) !== baseline, busy, 'Contrat')
  return <Box component="form" onSubmit={e => { e.preventDefault(); void save({ employee_type: kind, start_date: start, end_date: cdi ? null : end, is_cdi: cdi }).then(ok => { if (ok) setBaseline(JSON.stringify([kind, start, end, cdi])) }) }} sx={{ display: 'flex', flexWrap: 'wrap', gap: 2, alignItems: 'center' }}>
    <TextField select label="Type de contrat" value={kind} disabled={busy} onChange={e => { setKind(e.target.value); if (e.target.value === 'stagiaire') setCdi(false) }} sx={{ minWidth: 180 }}>{personnelTypes.map(([value, label]) => <MenuItem key={value} value={value}>{label}</MenuItem>)}</TextField>
    <TextField label="Début" type="date" required value={start} disabled={busy} onChange={e => setStart(e.target.value)} slotProps={{ inputLabel: { shrink: true } }} />
    <FormControlLabel control={<Checkbox checked={cdi} disabled={busy || kind === 'stagiaire'} onChange={e => setCdi(e.target.checked)} />} label="CDI / sans date de fin" />
    {!cdi && <TextField label="Fin" type="date" required value={end} disabled={busy} onChange={e => setEnd(e.target.value)} slotProps={{ inputLabel: { shrink: true }, htmlInput: { min: start || undefined } }} />}
    <Button type="submit" variant="contained" disabled={busy}>Enregistrer le contrat</Button>
  </Box>
}
