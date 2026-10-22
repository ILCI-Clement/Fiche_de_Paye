import { useState } from 'react'
import { Alert, Box, Button, Checkbox, CircularProgress, FormControlLabel, MenuItem, Paper, TextField, Typography } from '@mui/material'
import { useQuery } from '@tanstack/react-query'
import { adminApi } from '../admin/api'
import { dateLabel } from './api'
import type { EmploymentInput } from './api'
import { useRegistryViewer } from './context'
import { personDisplayName } from '../personDisplayName'

/** Copy only after an explicit choice; a contract start is not a proven entry date. */
export function ContractPrefill({ token, username, accountDisplayName, busy, onApply }: { token: string; username: string | null; accountDisplayName?: string | null; busy: boolean; onApply: (value: Partial<EmploymentInput>) => void }) {
  const [open, setOpen] = useState(false)
  const [selected, setSelected] = useState<number | ''>('')
  const [copyDates, setCopyDates] = useState(false)
  const [applied, setApplied] = useState(false)
  const viewer = useRegistryViewer()
  const contracts = useQuery({ queryKey: ['personnel-register', viewer, 'prefill', username], queryFn: ({ signal }) => adminApi.contracts(token, username!, signal), enabled: open && !!username, retry: false })
  const available = (contracts.data?.contracts || []).filter(item => !item.is_archived)
  const current = available.find(item => item.id === selected)
  const knownType = current?.is_cdi ? 'CDI' : current?.employee_type === 'stagiaire' ? 'stage' : ''
  function apply() {
    if (!current || busy) return
    onApply({ ...(knownType ? { contract_type: knownType, other_contract: '' } : {}), ...(copyDates ? { entry_date: current.start_date || null, exit_date: current.end_date || null } : {}) })
    setApplied(true)
  }
  return <Box sx={{ gridColumn: '1 / -1' }}>
    <Button size="small" variant="text" disabled={busy || !username} onClick={() => setOpen(value => !value)}>{open ? 'Masquer les contrats existants' : 'Consulter les contrats du compte associé'}</Button>
    {!username && <Typography variant="caption" sx={{ display: 'block' }} color="text.secondary">Associez un compte pour consulter ses contrats existants.</Typography>}
    {open && username && <Paper variant="outlined" sx={{ p: 2, mt: 1 }}>
      <Typography variant="body2" color="text.secondary" sx={{ mb: 2 }}>Source : contrats de {personDisplayName({ display_name: accountDisplayName })}. Le registre reste indépendant : aucune modification ne sera envoyée à la gestion des contrats ou aux fiches de présence.</Typography>
      {contracts.isPending ? <CircularProgress size={22} /> : contracts.error ? <Alert severity="error">{contracts.error.message}<Button onClick={() => void contracts.refetch()}>Réessayer</Button></Alert> : !available.length ? <Typography variant="body2">Aucun contrat non archivé associé à ce compte.</Typography> : <>
        <TextField select fullWidth label="Contrat à consulter" value={selected} disabled={busy} onChange={event => { setSelected(event.target.value === '' ? '' : Number(event.target.value)); setCopyDates(false); setApplied(false) }}>
          <MenuItem value="">Sélectionner un contrat</MenuItem>
          {available.map(item => <MenuItem key={item.id} value={item.id}>Contrat #{item.id} · {dateLabel(item.start_date)} — {item.is_cdi ? 'CDI' : dateLabel(item.end_date)}</MenuItem>)}
        </TextField>
        {current && <Box sx={{ mt: 2 }}>
          <Typography variant="body2">Type pouvant être repris : {knownType === 'CDI' ? 'CDI' : knownType === 'stage' ? 'Convention de stage' : 'indéterminé, à saisir manuellement'}.</Typography>
          <FormControlLabel sx={{ mt: 1 }} control={<Checkbox checked={copyDates} disabled={busy} onChange={event => { setCopyDates(event.target.checked); setApplied(false) }} />} label={<Typography variant="body2">J’ai vérifié que les dates de ce contrat correspondent à cette période d’emploi et je souhaite les reprendre.</Typography>} />
          <Typography variant="caption" sx={{ display: 'block' }} color="text.secondary">Un début de contrat n’est pas nécessairement la première date d’entrée dans l’entreprise. Une fin de contrat n’établit pas à elle seule une date de sortie.</Typography>
          <Button variant="outlined" disabled={busy || (!knownType && !copyDates)} onClick={apply} sx={{ mt: 1.5 }}>Préremplir les champs confirmés</Button>
          {applied && <Typography role="status" variant="body2" sx={{ mt: 1, color: 'primary.main' }}>Champs repris dans le formulaire. Vérifiez-les puis enregistrez pour confirmer.</Typography>}
        </Box>}
      </>}
    </Paper>}
  </Box>
}
