import { useDraftGuard } from '../draftGuard'
import { useRef, useState } from 'react'
import { Alert, Box, Button, Chip, Dialog, DialogActions, DialogContent, DialogTitle, MenuItem, TextField, Typography } from '@mui/material'
import { documentsApi, periodLabel } from './api'
import type { DocumentAccount } from './api'
import { personOptionLabel } from '../personDisplayName'

export function DocumentUpload({ token, username, currentDisplayName, admin, accounts, month, year, busy, run, requestError, onClose }: {
  token: string; username: string; currentDisplayName: string; admin: boolean; accounts: DocumentAccount[]; month: number; year: number; busy: boolean;
  run: (action: () => Promise<void>, message: string) => Promise<boolean>
  requestError: string; onClose: () => void;
}) {
  const [owner, setOwner] = useState(username)
  const people = accounts.some(person => person.username === username) ? accounts : [{ username, display_name: currentDisplayName }, ...accounts]
  const [type, setType] = useState('transport')
  const [other, setOther] = useState('')
  const [file, setFile] = useState<File | null>(null)
  const [error, setError] = useState('')
  const input = useRef<HTMLInputElement>(null)
  const close = useDraftGuard(!!file || !!other || owner !== username || type !== 'transport', busy, 'Document à déposer')
  async function upload() {
    if (!file) { setError('Choisissez un fichier.'); return }
    if (!file.size || file.size > 10 * 1024 * 1024) { setError('Le fichier doit contenir entre 1 octet et 10 Mo.'); return }
    if (!/\.(pdf|png|jpe?g)$/i.test(file.name)) { setError('Formats acceptés : PDF, JPG, PNG.'); return }
    if (type === 'autre' && !other.trim()) { setError('Précisez le type de document.'); return }
    setError('')
    const success = await run(async () => {
      const encoded = await new Promise<string>((resolve, reject) => {
        const reader = new FileReader()
        reader.onload = () => resolve(String(reader.result).split(',')[1])
        reader.onerror = () => reject(new Error('Impossible de lire le fichier.'))
        reader.readAsDataURL(file)
      })
      await documentsApi.upload(token, { employee_username: admin ? owner : username, original_filename: file.name, file_b64: encoded,
        document_type: type, other_type: other.trim(), receipt_month: month, receipt_year: year })
    }, 'Document enregistré.')
    if (success) onClose()
  }
  return <Dialog open fullWidth maxWidth="sm" aria-labelledby="document-upload-title" onClose={() => { if (!busy) close(onClose) }}>
    <DialogTitle id="document-upload-title">Déposer un document</DialogTitle>
    <DialogContent dividers>
      <Chip color="primary" variant="outlined" label={`Mois de rattachement : ${periodLabel(month, year)}`} sx={{ mb: 3 }} />
      <Box sx={{ display: 'flex', flexWrap: 'wrap', gap: 2, alignItems: 'center' }}>
        {admin && <TextField select label="Employé concerné" value={owner} disabled={busy} onChange={e => setOwner(e.target.value)} sx={{ minWidth: 220 }}>
          {people.map(person => <MenuItem key={person.username} value={person.username}>{personOptionLabel(person)}</MenuItem>)}
        </TextField>}
        <TextField select label="Type de document" value={type} disabled={busy} onChange={e => setType(e.target.value)} sx={{ minWidth: 180 }}>
          <MenuItem value="transport">Transport</MenuItem><MenuItem value="autre">Autre</MenuItem>
        </TextField>
        {type === 'autre' && <TextField label="Précisez le type" value={other} onChange={e => setOther(e.target.value)} disabled={busy} slotProps={{ htmlInput: { maxLength: 120 } }} />}
        <Button component="label" variant="outlined" disabled={busy}>Choisir un fichier<input ref={input} type="file" hidden accept=".pdf,.png,.jpg,.jpeg" onChange={e => { setFile(e.target.files?.[0] || null); setError('') }} /></Button>
        <Typography variant="caption" color="text.secondary">PDF, JPG, PNG · 10 Mo maximum</Typography>
      </Box>
      <Box sx={{ display: 'flex', flexWrap: 'wrap', gap: 2, alignItems: 'center', mt: 2 }}>
        <Typography sx={{ overflowWrap: 'anywhere' }}>{file?.name || 'Aucun fichier sélectionné'}</Typography>
      </Box>
      {error && <Alert severity="error" sx={{ mt: 2 }}>{error}</Alert>}
      {requestError && <Alert severity="error" sx={{ mt: 2 }}>{requestError}</Alert>}
    </DialogContent>
    <DialogActions sx={{ p: 2 }}>
      <Button disabled={busy} onClick={() => close(onClose)}>Annuler</Button>
      <Button variant="contained" disabled={busy || !file} onClick={() => void upload()}>{busy ? 'Envoi…' : 'Envoyer le document'}</Button>
    </DialogActions>
  </Dialog>
}
