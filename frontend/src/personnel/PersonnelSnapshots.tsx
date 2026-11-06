import { useRef, useState } from 'react'
import { useQuery, useQueryClient } from '@tanstack/react-query'
import { Alert, Box, Button, Checkbox, CircularProgress, Dialog, DialogActions, DialogContent, DialogTitle, FormControlLabel, IconButton, MenuItem, Pagination, Stack, TextField, Tooltip, Typography } from '@mui/material'
import DownloadOutlined from '@mui/icons-material/DownloadOutlined'
import { useDraftGuard } from '../draftGuard'
import { downloadFile } from '../download'
import { snapshotApi, type RegisterSnapshot } from './snapshotApi'

const time = (value: string | null) => value ? new Intl.DateTimeFormat('fr-FR', { dateStyle: 'short', timeStyle: 'short', timeZone: 'Europe/Paris' }).format(new Date(value.endsWith('Z') ? value : value + 'Z')) : '—'
const errorLabel = (row: RegisterSnapshot) => row.error_code === 'row_limit' ? 'Trop de lignes : aucune archive partielle créée.' : 'Génération impossible. Vérifiez les clés de chiffrement et le service, puis réessayez.'

export function PersonnelSnapshots({ token, viewer, onClose }: { token: string; viewer: string; onClose: () => void }) {
  const client = useQueryClient()
  const [page, setPage] = useState(1)
  const [cycle, setCycle] = useState<number | null>(null)
  const [busy, setBusy] = useState(false)
  const [error, setError] = useState('')
  const [message, setMessage] = useState('')
  const [download, setDownload] = useState<RegisterSnapshot | null>(null)
  const [includeNir, setIncludeNir] = useState(false)
  const [format, setFormat] = useState<'pdf' | 'xlsx'>('pdf')
  const lock = useRef(false)
  const requestId = useRef<string | null>(null)
  const cycleVersion = useRef<number | null>(null)
  const query = useQuery({ queryKey: ['personnel-snapshots', viewer, page], queryFn: ({ signal }) => snapshotApi.list(token, page, signal), retry: false, gcTime: 0, refetchInterval: busy ? false : 30000 })
  const data = query.data
  const changed = cycle !== null && cycle !== data?.settings.interval_months
  const close = useDraftGuard(changed, busy, 'Cycle des archives du registre')
  async function run(action: () => Promise<unknown>, success: string) {
    if (lock.current) return
    lock.current = true; setBusy(true); setError(''); setMessage('')
    try { await action(); setMessage(success) }
    catch (failure) { setError(failure instanceof Error ? failure.message : 'Action impossible.') }
    finally {
      await Promise.allSettled([client.invalidateQueries({ queryKey: ['personnel-snapshots'] }), client.invalidateQueries({ queryKey: ['work-center'] })])
      lock.current = false; setBusy(false)
    }
  }
  return <>
    <Dialog open fullWidth maxWidth="lg" onClose={() => { if (!busy) close(onClose) }} aria-labelledby="register-snapshots-title">
      <DialogTitle id="register-snapshots-title">Historique des archives</DialogTitle>
      <DialogContent dividers>
        <Typography variant="body2" color="text.secondary" sx={{ mb: 2 }}>Personnel non archivé, toutes les périodes d’emploi. Fichiers historiques chiffrés, conservés sans remplacement. Heures de Paris.</Typography>
        {error && <Alert severity="error" sx={{ mb: 1 }}>{error}</Alert>}
        {message && <Alert severity="info" sx={{ mb: 1 }}>{message}</Alert>}
        {query.isPending ? <CircularProgress size={24} /> : query.error ? <Alert severity="error">{query.error.message}<Button onClick={() => void query.refetch()}>Réessayer</Button></Alert> : data && <>
          {!data.encryption_ready && <Alert severity="warning" sx={{ mb: 2 }}>Clé des archives indisponible. La génération et le téléchargement nécessitent une clé dédiée.</Alert>}
          <Box sx={{ display: 'flex', alignItems: 'center', gap: 1.5, flexWrap: 'wrap', mb: 2 }}>
            <TextField select label="Cycle automatique" value={cycle ?? data.settings.interval_months} disabled={busy || !data.can_configure} sx={{ width: 180 }} onChange={event => { cycleVersion.current ??= data.settings.version; setCycle(Number(event.target.value)) }}>
              {[1, 3, 6, 12].map(months => <MenuItem key={months} value={months}>Tous les {months} mois</MenuItem>)}
            </TextField>
            {data.can_configure && changed && <Button disabled={busy} onClick={() => void run(async () => { await snapshotApi.configure(token, cycle!, cycleVersion.current!); setCycle(null); cycleVersion.current = null }, 'Cycle mis à jour.')}>Enregistrer le cycle</Button>}
            <Typography variant="caption" color="text.secondary">Prochaine : {time(data.settings.next_due)}</Typography>
            <Button variant="contained" disabled={busy || !data.encryption_ready} sx={{ ml: 'auto' }} onClick={() => void run(async () => {
              requestId.current ||= crypto.randomUUID()
              const result = await snapshotApi.manual(token, requestId.current)
              requestId.current = null
              if (result.status === 'failed') throw new Error(errorLabel(result))
              setPage(1)
            }, 'Archive créée.')}>{busy ? 'Traitement…' : 'Archiver maintenant'}</Button>
            <Button disabled={busy} onClick={() => void query.refetch()}>Actualiser</Button>
          </Box>
          <Typography variant="caption" color="text.secondary">Le cycle suit le calendrier à 03:00. Une archive manuelle ne modifie pas la prochaine échéance. Un rattrapage reflète la date réelle de collecte.</Typography>
          {!data.snapshots.length && <Typography sx={{ py: 4 }}>Aucune archive créée. La première sera générée à l’échéance ou manuellement.</Typography>}
          {data.snapshots.map(row => <Box key={row.id} sx={{ display: 'flex', alignItems: 'center', gap: 2, py: 1.5, borderBottom: 1, borderColor: 'divider' }}>
            <Box sx={{ minWidth: 0, flex: 1 }}>
              <Typography variant="body2" sx={{ fontWeight: 600 }}>{row.source === 'manual' ? 'Manuelle' : row.source === 'catchup' ? 'Rattrapage' : 'Automatique'} · {time(row.captured_at || row.created_at)}</Typography>
              {row.scheduled_for && <Typography variant="caption" color="text.secondary">Planifiée : {time(row.scheduled_for)} · Collectée : {time(row.captured_at)}</Typography>}
              <Typography variant="caption" color="text.secondary" sx={{ display: 'block' }}>{row.status === 'ready' ? `${row.employee_count} personne(s) · ${row.row_count} ligne(s)` : row.status === 'failed' ? errorLabel(row) : 'Génération en cours'}</Typography>
            </Box>
            {row.status === 'ready' ? <Tooltip title="Télécharger l’archive"><span><IconButton aria-label="Télécharger l’archive" disabled={busy || !data.encryption_ready} onClick={() => { setIncludeNir(false); setFormat('pdf'); setDownload(row) }}><DownloadOutlined /></IconButton></span></Tooltip> : row.status === 'failed' && <Button disabled={busy || !data.encryption_ready} onClick={() => void run(async () => { const result = await snapshotApi.retry(token, row.id); if (result.status === 'failed') throw new Error(errorLabel(result)) }, 'Archive générée à la date actuelle.')}>Réessayer</Button>}
          </Box>)}
          {data.total > 20 && <Pagination sx={{ mt: 2 }} count={Math.ceil(data.total / 20)} page={page} disabled={busy} onChange={(_, next) => setPage(next)} />}
        </>}
      </DialogContent>
      <DialogActions><Button disabled={busy} onClick={() => close(onClose)}>Fermer</Button></DialogActions>
    </Dialog>
    <Dialog open={!!download} fullWidth maxWidth="sm" onClose={() => { if (!busy) { setDownload(null); setIncludeNir(false) } }}>
      <DialogTitle>Télécharger l’archive</DialogTitle><DialogContent>
        <Stack spacing={1}>
          {data?.can_export_excel ? <TextField select label="Format" size="small" value={format} disabled={busy} onChange={event => setFormat(event.target.value as 'pdf' | 'xlsx')}><MenuItem value="pdf">PDF</MenuItem><MenuItem value="xlsx">Excel</MenuItem></TextField> : <Typography variant="body2">Format : PDF</Typography>}
          <Typography variant="body2">État historique du {time(download?.captured_at || null)}. Le téléchargement est enregistré.</Typography>
          <FormControlLabel control={<Checkbox checked={includeNir} disabled={busy} onChange={event => setIncludeNir(event.target.checked)} />} label="Inclure les numéros de sécurité sociale complets" />
          {includeNir && <Alert severity="warning">Le fichier téléchargé contiendra les numéros en clair.</Alert>}
          {error && <Alert severity="error">{error}</Alert>}
        </Stack>
      </DialogContent><DialogActions><Button disabled={busy} onClick={() => { setDownload(null); setIncludeNir(false) }}>Annuler</Button>
        <Button variant="contained" disabled={busy || !download} onClick={() => void run(async () => { const result = await snapshotApi.download(token, download!.id, includeNir, data?.can_export_excel ? format : 'pdf'); downloadFile(result); setDownload(null); setIncludeNir(false) }, 'Archive téléchargée.')}>Télécharger</Button>
      </DialogActions>
    </Dialog>
  </>
}
