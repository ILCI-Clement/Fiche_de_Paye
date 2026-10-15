import { useDraftGuard } from '../draftGuard'
import { useRef, useState } from 'react'
import { Accordion, AccordionDetails, AccordionSummary, Alert, Box, Button, Chip, MenuItem, TextField, Typography } from '@mui/material'
import ExpandMoreRounded from '@mui/icons-material/ExpandMoreRounded'
import { documentsApi, periodLabel } from './api'

export function DocumentUpload({ token, username, admin, users, month, year, busy, run }: {
  token: string; username: string; admin: boolean; users: string[]; month: number; year: number; busy: boolean;
  run: (action: () => Promise<void>, message: string) => Promise<boolean>
}) {
  const [owner, setOwner] = useState(username)
  const [type, setType] = useState('transport')
  const [other, setOther] = useState('')
  const [file, setFile] = useState<File | null>(null)
  const [error, setError] = useState('')
  const input = useRef<HTMLInputElement>(null)
  useDraftGuard(!!file || !!other, busy, 'Document à déposer')
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
    if (success) { setFile(null); setOther(''); if (input.current) input.current.value = '' }
  }
  return <Accordion disableGutters>
    <AccordionSummary expandIcon={<ExpandMoreRounded />}><Typography sx={{ fontWeight: 650 }}>Déposer un document</Typography></AccordionSummary>
    <AccordionDetails>
      <Box sx={{ display: 'flex', flexWrap: 'wrap', gap: 2, alignItems: 'center' }}>
        {admin && <TextField select label="Employé concerné" value={owner} disabled={busy} onChange={e => setOwner(e.target.value)} sx={{ minWidth: 220 }}>
          {[...new Set([username, ...users])].map(user => <MenuItem key={user} value={user}>{user.split('@')[0]}</MenuItem>)}
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
        <Chip color="primary" variant="outlined" label={`Mois de rattachement : ${periodLabel(month, year)}`} />
        <Button variant="contained" disabled={busy || !file} onClick={() => void upload()}>Envoyer le document</Button>
      </Box>
      {error && <Alert severity="error" sx={{ mt: 2 }}>{error}</Alert>}
    </AccordionDetails>
  </Accordion>
}
