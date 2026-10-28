import { useRef, useState } from 'react'
import { Alert, Box, Button, Chip, Dialog, DialogActions, DialogContent, DialogTitle, Divider, IconButton, TextField, Tooltip, Typography } from '@mui/material'
import EditOutlined from '@mui/icons-material/EditOutlined'
import ArchiveOutlined from '@mui/icons-material/ArchiveOutlined'
import UnarchiveOutlined from '@mui/icons-material/UnarchiveOutlined'
import { useQueryClient } from '@tanstack/react-query'
import { useDraftGuard } from '../draftGuard'
import { categoryApi } from './api'
import type { DocumentCategory } from './api'
import { DocumentTypeBadge } from './DocumentTypeBadge'

export function CategoryManager({ token, categories, onClose }: { token: string; categories: DocumentCategory[]; onClose: () => void }) {
  const client = useQueryClient()
  const [editing, setEditing] = useState<DocumentCategory | null>(null)
  const [label, setLabel] = useState('')
  const [busy, setBusy] = useState(false)
  const [error, setError] = useState('')
  const lock = useRef(false)
  const dirty = editing ? label !== editing.label : !!label.trim()
  const close = useDraftGuard(dirty, busy, 'Types de documents')
  async function perform(action: () => Promise<unknown>, reset = false) {
    if (lock.current) return
    lock.current = true; setBusy(true); setError('')
    try {
      await action()
      if (reset) { setLabel(''); setEditing(null) }
      await Promise.all([client.invalidateQueries({ queryKey: ['document-categories'] }), client.invalidateQueries({ queryKey: ['documents'] })])
    } catch (failure) { setError(failure instanceof Error ? failure.message : 'Modification impossible.') }
    finally { lock.current = false; setBusy(false) }
  }
  return <Dialog open fullWidth maxWidth="sm" onClose={() => { if (!busy) close(onClose) }}>
    <DialogTitle>Types de documents</DialogTitle>
    <DialogContent dividers>
      <Box component="form" onSubmit={event => { event.preventDefault(); void perform(() => editing ? categoryApi.update(token, editing, { label: label.trim() }) : categoryApi.create(token, label.trim()), true) }} sx={{ display: 'flex', alignItems: 'center', gap: 1, flexWrap: 'wrap', mb: 2 }}>
        <TextField autoFocus required size="small" label={editing ? 'Modifier le libellé' : 'Nouveau type'} value={label} disabled={busy} onChange={event => setLabel(event.target.value)} slotProps={{ htmlInput: { maxLength: 80 } }} sx={{ flex: '1 1 200px' }} />
        <Button type="submit" variant="contained" disabled={busy || !label.trim() || (!!editing && !dirty)}>{editing ? 'Enregistrer' : 'Ajouter'}</Button>
        {editing && <Button disabled={busy} onClick={() => { setLabel(''); setEditing(null); setError('') }}>Annuler</Button>}
      </Box>
      {error && <Alert severity="error" sx={{ mb: 1.5 }}>{error}<Button color="inherit" size="small" disabled={busy} onClick={() => { setEditing(null); setLabel(''); void client.invalidateQueries({ queryKey: ['document-categories'] }) }}>Actualiser</Button></Alert>}
      {categories.map(category => <Box key={category.code}><Box sx={{ display: 'flex', alignItems: 'center', gap: 1, py: 1, minWidth: 0 }}>
        <Box sx={{ flex: 1, minWidth: 0 }}><DocumentTypeBadge code={category.code} label={category.label} inactive={!category.is_active} /></Box>
        {category.is_system ? <Chip size="small" label="Standard" variant="outlined" /> : <>
          {!category.is_active && <Chip size="small" label="Désactivé" variant="outlined" />}
          <Tooltip title={`Renommer ${category.label}`}><span><IconButton size="small" aria-label={`Renommer ${category.label}`} disabled={busy} onClick={() => close(() => { setEditing(category); setLabel(category.label); setError('') })}><EditOutlined fontSize="small" /></IconButton></span></Tooltip>
          <Tooltip title={category.is_active ? 'Désactiver pour les nouveaux dépôts' : 'Réactiver'}><span><IconButton size="small" aria-label={`${category.is_active ? 'Désactiver' : 'Réactiver'} ${category.label}`} disabled={busy} onClick={() => void perform(() => categoryApi.update(token, category, { is_active: !category.is_active }))}>{category.is_active ? <ArchiveOutlined fontSize="small" /> : <UnarchiveOutlined fontSize="small" />}</IconButton></span></Tooltip>
        </>}
      </Box><Divider /></Box>)}
      <Typography variant="caption" color="text.secondary" sx={{ display: 'block', mt: 1.5 }}>Les types désactivés restent visibles sur les documents existants.</Typography>
    </DialogContent>
    <DialogActions><Button disabled={busy} onClick={() => close(onClose)}>Fermer</Button></DialogActions>
  </Dialog>
}
