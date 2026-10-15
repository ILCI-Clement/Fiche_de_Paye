import { useDraftGuard, useWorkspaceGuard } from '../draftGuard'
import { useState } from 'react'
import { Box, Button, Checkbox, FormControlLabel, MenuItem, TextField } from '@mui/material'
import type { Group } from '../team/api'
import { adminApi } from './api'
import type { RunAction } from './api'

export function GroupEditor({ token, groups, busy, run }: { token: string; groups: Group[]; busy: boolean; run: RunAction }) {
  const guard = useWorkspaceGuard()
  const [selected, setSelected] = useState<number | ''>('')
  const current = groups.find(g => g.id === selected)
  return <>
    <TextField select label="Département à gérer" value={selected} disabled={busy} fullWidth onChange={e => { const value = e.target.value === '' ? '' : Number(e.target.value); guard?.request(() => setSelected(value)) }}><MenuItem value="">Nouveau département</MenuItem>{groups.map(g => <MenuItem key={g.id} value={g.id}>{g.name}{!g.is_active ? ' · inactif' : ''}</MenuItem>)}</TextField>
    <GroupForm key={`${selected}-${JSON.stringify(current)}`} current={current} busy={busy} save={(name, active) => run(() => adminApi.saveGroup(token, current?.id ?? null, name, active), 'Département enregistré.')} />
  </>
}

function GroupForm({ current, busy, save }: { current?: Group; busy: boolean; save: (name: string, active: boolean) => Promise<boolean> }) {
  const [name, setName] = useState(current?.name || '')
  const [active, setActive] = useState(current ? Boolean(current.is_active) : true)
  const [baseline, setBaseline] = useState(() => JSON.stringify([name, active]))
  useDraftGuard(JSON.stringify([name, active]) !== baseline, busy, 'Département')
  return <Box component="form" onSubmit={e => { e.preventDefault(); if (name.trim()) void save(name.trim(), active).then(ok => { if (ok) { setName(current?.id ? name.trim() : ''); setBaseline(JSON.stringify([current?.id ? name.trim() : '', active])) } }) }} sx={{ display: 'flex', flexWrap: 'wrap', alignItems: 'center', gap: 2, mt: 2 }}>
    <TextField label="Nom du département" required value={name} disabled={busy} onChange={e => setName(e.target.value)} sx={{ flex: 1, minWidth: 200 }} />
    <FormControlLabel control={<Checkbox checked={active} disabled={busy} onChange={e => setActive(e.target.checked)} />} label="Actif" />
    <Button type="submit" variant="contained" disabled={busy}>Enregistrer le département</Button>
  </Box>
}
