import { useDraftGuard } from '../draftGuard'
import { useState, type ReactNode } from 'react'
import { Box, Button, Checkbox, Dialog, DialogActions, DialogContent, DialogTitle, FormControlLabel, MenuItem, TextField } from '@mui/material'
import type { Group } from '../team/api'
import { adminApi } from './api'
import type { RunAction } from './api'

export function GroupEditor({ token, groups, busy, run, onClose }: { token: string; groups: Group[]; busy: boolean; run: RunAction; onClose: () => void }) {
  const [selected, setSelected] = useState<number | ''>('')
  const current = groups.find(g => g.id === selected)
  const selector = (request: (action: () => void) => void) =>
    <TextField select label="Département à gérer" value={selected} disabled={busy} fullWidth onChange={e => { const value = e.target.value === '' ? '' : Number(e.target.value); request(() => setSelected(value)) }}><MenuItem value="">Nouveau département</MenuItem>{groups.map(g => <MenuItem key={g.id} value={g.id}>{g.name}{!g.is_active ? ' · inactif' : ''}</MenuItem>)}</TextField>
  return <GroupForm key={`${selected}-${JSON.stringify(current)}`} selector={selector} onClose={onClose} current={current} busy={busy} save={(name, active) => run(() => adminApi.saveGroup(token, current?.id ?? null, name, active), 'Département enregistré.')} />
}

function GroupForm({ current, busy, save, selector, onClose }: { current?: Group; busy: boolean; save: (name: string, active: boolean) => Promise<boolean>; selector: (request: (action: () => void) => void) => ReactNode; onClose: () => void }) {
  const [name, setName] = useState(current?.name || '')
  const [active, setActive] = useState(current ? Boolean(current.is_active) : true)
  const [baseline, setBaseline] = useState(() => JSON.stringify([name, active]))
  const close = useDraftGuard(JSON.stringify([name, active]) !== baseline, busy, 'Département')
  return <Dialog open fullWidth maxWidth="sm" onClose={() => { if (!busy) close(onClose) }}>
    <DialogTitle>Gérer les départements</DialogTitle><DialogContent sx={{ pt: '12px !important' }}>{selector(close)}
    <Box component="form" onSubmit={e => { e.preventDefault(); if (name.trim()) void save(name.trim(), active).then(ok => { if (ok) { setName(current?.id ? name.trim() : ''); setBaseline(JSON.stringify([current?.id ? name.trim() : '', active])) } }) }} sx={{ display: 'flex', flexWrap: 'wrap', alignItems: 'center', gap: 2, mt: 2 }}>
    <TextField label="Nom du département" required value={name} disabled={busy} onChange={e => setName(e.target.value)} sx={{ flex: 1, minWidth: 200 }} />
    <FormControlLabel control={<Checkbox checked={active} disabled={busy} onChange={e => setActive(e.target.checked)} />} label="Actif" />
    <Button type="submit" variant="contained" disabled={busy}>Enregistrer le département</Button>
  </Box></DialogContent><DialogActions><Button disabled={busy} onClick={() => close(onClose)}>Fermer</Button></DialogActions></Dialog>
}
