import { useEffect, useRef, useState } from 'react'
import { Alert, Button, Checkbox, Dialog, DialogActions, DialogContent, DialogTitle, FormControlLabel, MenuItem, TextField, Typography } from '@mui/material'
import DownloadOutlined from '@mui/icons-material/DownloadOutlined'
import { useQueryClient } from '@tanstack/react-query'
import { useDraftGuard } from '../draftGuard'
import { downloadFile } from '../download'
import { personnelApi } from './api'
import type { PersonnelFilters } from './api'

export function PersonnelExport({ token, filters, total, canExportExcel, onClose, onSuccess }: {
  token: string; filters: PersonnelFilters; total: number; canExportExcel: boolean; onClose: () => void; onSuccess: () => void;
}) {
  // The dialog owns the chosen filter snapshot; background refresh cannot change it.
  const [selection] = useState(() => ({ ...filters }))
  const [count] = useState(total)
  const [confirmed, setConfirmed] = useState(false)
  const [format, setFormat] = useState<'pdf' | 'xlsx'>('pdf')
  const [busy, setBusy] = useState(false)
  const [error, setError] = useState('')
  const pending = useRef<AbortController | null>(null)
  const client = useQueryClient()
  useDraftGuard(false, busy, 'Export des utilisateurs')
  useEffect(() => () => pending.current?.abort(), [])
  async function exportExcel() {
    if (!confirmed || pending.current) return
    const controller = new AbortController()
    pending.current = controller; setBusy(true); setError('')
    try {
      const file = await personnelApi.exportFile(token, selection, controller.signal, canExportExcel ? format : 'pdf')
      if (controller.signal.aborted) return
      downloadFile(file)
      void client.invalidateQueries({ queryKey: ['personnel-register'] })
      onSuccess(); onClose()
    } catch (reason) {
      if (!controller.signal.aborted) setError(reason instanceof Error ? reason.message : 'Impossible de générer le fichier.')
    } finally {
      pending.current = null
      if (!controller.signal.aborted) setBusy(false)
    }
  }
  return <Dialog open fullWidth maxWidth="sm" aria-labelledby="personnel-export-title" onClose={() => { if (!busy) onClose() }}>
    <DialogTitle id="personnel-export-title">Exporter les utilisateurs</DialogTitle>
    <DialogContent dividers>
      {canExportExcel ? <TextField select fullWidth size="small" label="Format" value={format} disabled={busy} onChange={event => setFormat(event.target.value as 'pdf' | 'xlsx')} sx={{ mb: 2 }}><MenuItem value="pdf">PDF</MenuItem><MenuItem value="xlsx">Excel</MenuItem></TextField> : <Typography sx={{ mb: 2 }}>Format : PDF</Typography>}
      <Typography variant="body2" sx={{ mb: 2 }}>{count} personne(s) correspondent aux filtres actuels, toutes les pages incluses. Une ligne sera créée pour chaque période d’emploi. Les champs non renseignés resteront vides.</Typography>
      <Alert severity="warning" sx={{ mb: 2 }}>Le fichier contiendra des données personnelles et les numéros de sécurité sociale en clair. Conservez-le dans un emplacement protégé. L’export sera enregistré dans l’historique des personnes concernées.</Alert>
      {error && <Alert severity="error" sx={{ mb: 2 }}>{error}</Alert>}
      <FormControlLabel control={<Checkbox checked={confirmed} disabled={busy} onChange={event => setConfirmed(event.target.checked)} />} label="Je confirme l’export de ces données personnelles et des numéros de sécurité sociale." />
    </DialogContent>
    <DialogActions><Button disabled={busy} onClick={onClose}>Annuler</Button><Button variant="contained" startIcon={<DownloadOutlined />} disabled={busy || !confirmed} onClick={() => void exportExcel()}>{busy ? 'Génération…' : `Télécharger le ${format === 'xlsx' && canExportExcel ? 'fichier Excel' : 'PDF'}`}</Button></DialogActions>
  </Dialog>
}
