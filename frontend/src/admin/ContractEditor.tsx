import { FrenchDateField } from '../components/FrenchDateField'
import { useDraftGuard, useWorkspaceGuard } from '../draftGuard'
import { useState } from 'react'
import { Alert, Box, Button, Checkbox, CircularProgress, Dialog, DialogActions, DialogContent, DialogTitle, FormControlLabel, MenuItem, TextField, ToggleButton, ToggleButtonGroup, Typography } from '@mui/material'
import { useQuery } from '@tanstack/react-query'
import { adminApi, personnelTypes } from './api'
import type { Contract, ContractInput, RunAction } from './api'
import { DeleteAction } from '../components/DeleteAction'
import { PermanentDeleteDialog } from '../components/PermanentDeleteDialog'

export function ContractEditor({ token, username, busy, run, canDelete = false, readOnly = false, requestError = '' }: { token: string; username: string; busy: boolean; run: RunAction; canDelete?: boolean; readOnly?: boolean; requestError?: string }) {
  const contracts = useQuery({ queryKey: ['admin-contracts', username], queryFn: ({ signal }) => adminApi.contracts(token, username, signal), retry: false })
  const guard = useWorkspaceGuard()
  const protect = (action: () => void) => guard ? guard.request(action) : action()
  const [selected, setSelected] = useState<number | ''>('')
  const [archived, setArchived] = useState(false)
  const [archiveConfirm, setArchiveConfirm] = useState(false)
  const [deleting, setDeleting] = useState(false)
  const [failed, setFailed] = useState(false)
  if (contracts.isPending) return <CircularProgress size={24} />
  if (contracts.error) return <Alert severity="error">{contracts.error.message}<Button onClick={() => void contracts.refetch()}>Réessayer</Button></Alert>
  const current = contracts.data.contracts.find(c => c.id === selected)
  const disabled = busy || readOnly
  const options = contracts.data.contracts.filter(contract => !!contract.is_archived === archived)
  return <>
    <ToggleButtonGroup exclusive size="small" value={archived ? 'archived' : 'active'} disabled={busy} sx={{ mb: 2 }} aria-label="État des contrats" onChange={(_, value) => { if (value) protect(() => { setArchived(value === 'archived'); setSelected('') }) }}><ToggleButton value="active">Normal</ToggleButton><ToggleButton value="archived">Archivé</ToggleButton></ToggleButtonGroup>
    <TextField select fullWidth label="Contrat à gérer" value={selected} disabled={busy} onChange={e => { const value = e.target.value === '' ? '' : Number(e.target.value); protect(() => setSelected(value)) }}>
      <MenuItem value="">{archived ? 'Sélectionner un contrat archivé' : 'Nouveau contrat'}</MenuItem>{options.map(c => <MenuItem key={c.id} value={c.id}>{c.employee_type} · {c.start_date} → {c.end_date || 'CDI'}</MenuItem>)}
    </TextField>
    <Typography variant="body2" color="text.secondary" sx={{ my: 2 }}>Les contrats historiques sont conservés. Les périodes ne doivent pas se chevaucher.</Typography>
    {(!archived || current) && <ContractForm key={`${selected}-${JSON.stringify(current)}`} current={current} busy={busy} disabled={disabled || !!current?.is_archived} save={data => disabled || current?.is_archived ? Promise.resolve(false) : run(() => adminApi.saveContract(token, username, current?.id ?? null, data), 'Contrat enregistré.')} />}
    {current && !readOnly && <Box sx={{ display: 'flex', alignItems: 'center', gap: 1, mt: 2 }}>
      <Button disabled={busy} onClick={() => protect(() => { setFailed(false); setArchiveConfirm(true) })}>{current.is_archived ? 'Restaurer le contrat' : 'Archiver le contrat'}</Button>
      {canDelete && current.is_archived && <DeleteAction label="Supprimer définitivement ce contrat" disabled={busy} onClick={() => protect(() => setDeleting(true))} />}
    </Box>}
    {current && <Dialog open={archiveConfirm} onClose={() => { if (!busy) setArchiveConfirm(false) }}><DialogTitle>{current.is_archived ? 'Restaurer ce contrat ?' : 'Archiver ce contrat ?'}</DialogTitle><DialogContent><Alert severity="info">Les instantanés des fiches de présence et l’historique sont conservés. Un contrat archivé n’est plus proposé pour de nouvelles fiches.</Alert>{failed && <Alert severity="error" sx={{ mt: 2 }}>{requestError || 'Opération refusée. Actualisez les données avant de réessayer.'}</Alert>}</DialogContent><DialogActions><Button disabled={busy} onClick={() => setArchiveConfirm(false)}>Annuler</Button><Button disabled={disabled} onClick={async () => { setFailed(false); const ok = await run(() => adminApi.archiveContract(token, username, current.id, !current.is_archived), current.is_archived ? 'Contrat restauré.' : 'Contrat archivé.'); if (ok) { setSelected(''); setArchiveConfirm(false) } else setFailed(true) }}>Confirmer</Button></DialogActions></Dialog>}
    {current && deleting && <PermanentDeleteDialog title="Supprimer définitivement ce contrat ?" expected={String(current.id)} description="Le contrat archivé sera supprimé uniquement s’il n’est référencé par aucune fiche de présence." onConfirm={confirmation => adminApi.deleteContract(token, username, current.id, confirmation)} onClose={() => setDeleting(false)} onDeleted={() => { setDeleting(false); setSelected(''); void run(async () => {}, 'Contrat supprimé.') }} />}
  </>
}

function ContractForm({ current, busy, disabled, save }: { current?: Contract; busy: boolean; disabled: boolean; save: (data: ContractInput) => Promise<boolean> }) {
  const [kind, setKind] = useState(current?.employee_type || 'salarie')
  const [start, setStart] = useState(current?.start_date || '')
  const [end, setEnd] = useState(current?.end_date || '')
  const [cdi, setCdi] = useState(Boolean(current?.is_cdi))
  const [baseline, setBaseline] = useState(() => JSON.stringify([kind, start, end, cdi]))
  useDraftGuard(JSON.stringify([kind, start, end, cdi]) !== baseline, busy, 'Contrat')
  return <Box component="form" onSubmit={e => { e.preventDefault(); void save({ employee_type: kind, start_date: start, end_date: cdi ? null : end, is_cdi: cdi }).then(ok => { if (ok) setBaseline(JSON.stringify([kind, start, end, cdi])) }) }} sx={{ display: 'flex', flexWrap: 'wrap', gap: 2, alignItems: 'center' }}>
    <TextField select label="Type de contrat" value={kind} disabled={disabled} onChange={e => { setKind(e.target.value); if (e.target.value === 'stagiaire') setCdi(false) }} sx={{ minWidth: 180 }}>{personnelTypes.map(([value, label]) => <MenuItem key={value} value={value}>{label}</MenuItem>)}</TextField>
    <FrenchDateField label="Début" required value={start} disabled={disabled} onChange={setStart} />
    <FormControlLabel control={<Checkbox checked={cdi} disabled={disabled || kind === 'stagiaire'} onChange={e => setCdi(e.target.checked)} />} label="CDI / sans date de fin" />
    {!cdi && <FrenchDateField label="Fin" required value={end} disabled={disabled} onChange={setEnd} min={start || undefined} />}
    <Button type="submit" variant="contained" disabled={disabled}>Enregistrer le contrat</Button>
  </Box>
}
