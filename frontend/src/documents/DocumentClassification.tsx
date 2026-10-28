import { useState } from 'react'
import { Alert, Box, Button, Dialog, DialogActions, DialogContent, DialogTitle, Typography } from '@mui/material'
import { useDraftGuard } from '../draftGuard'
import type { DocumentCategory, Receipt } from './api'
import { DocumentTypeFields } from './DocumentTypeFields'

export function DocumentClassification({ receipt, categories, busy, error, onSave, onClose }: {
  receipt: Receipt; categories: DocumentCategory[]; busy: boolean; error: string;
  onSave: (type: string, other: string) => Promise<boolean>; onClose: () => void
}) {
  const [type, setType] = useState(receipt.document_type || '')
  const [other, setOther] = useState(receipt.other_type || '')
  const dirty = type !== (receipt.document_type || '') || (type === 'autre' ? other.trim() : null) !== receipt.other_type
  const close = useDraftGuard(dirty, busy, 'Type du document')
  return <Dialog open fullWidth maxWidth="sm" onClose={() => { if (!busy) close(onClose) }}>
    <DialogTitle>Modifier le type</DialogTitle><DialogContent dividers>
      <Typography variant="body2" sx={{ mb: 2, overflowWrap: 'anywhere' }}>{receipt.original_filename}</Typography>
      <Box sx={{ display: 'flex', flexWrap: 'wrap', gap: 2 }}><DocumentTypeFields categories={categories} type={type} other={other} onType={setType} onOther={setOther} disabled={busy} /></Box>
      {error && <Alert severity="error" sx={{ mt: 2 }}>{error}</Alert>}
    </DialogContent><DialogActions>
      <Button disabled={busy} onClick={() => close(onClose)}>Annuler</Button>
      <Button variant="contained" disabled={busy || !dirty || !categories.some(category => category.code === type && category.is_active) || (type === 'autre' && !other.trim())} onClick={async () => { if (await onSave(type, other.trim())) onClose() }}>Enregistrer</Button>
    </DialogActions>
  </Dialog>
}
