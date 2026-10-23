import { useCallback, useRef, useState } from 'react'
import { Alert, Box, Button, Chip, CircularProgress, IconButton, Pagination, Paper, ToggleButton, ToggleButtonGroup, Tooltip, Typography } from '@mui/material'
import DownloadOutlined from '@mui/icons-material/DownloadOutlined'
import VisibilityOutlined from '@mui/icons-material/VisibilityOutlined'
import RefreshOutlined from '@mui/icons-material/RefreshOutlined'
import { useQuery } from '@tanstack/react-query'
import { signedFichesApi, signedStatus } from './signedFichesApi'
import type { SignedFiche } from './signedFichesApi'
import { DocumentPreview } from '../documents/DocumentPreview'
import { downloadFile } from '../download'
import { periodLabel } from '../documents/api'
import { useDraftGuard } from '../draftGuard'
import { useWorkspaceToolbar } from '../workspaceToolbar'
import { useCapabilities } from '../capabilities'

function SignedPreview({ item, token, onClose }: { item: SignedFiche; token: string; onClose: () => void }) {
  const loadFile = useCallback((signal: AbortSignal) => signedFichesApi.file(token, item.id, signal), [token, item.id])
  return <DocumentPreview token={token} receipt={{ id: 0, original_filename: `${item.employee_name} · ${periodLabel(item.fiche_month, item.fiche_year)} · signé.pdf` }} loadFile={loadFile} onClose={onClose} />
}

export function SignedFichesPanel({ token, month, year, username, viewer }: { token: string; month: number; year: number; username?: string; viewer: string }) {
  const [selectedMode, setMode] = useState('production')
  const localSandbox = useCapabilities().data?.environment === 'local-snapshot'
  const mode = localSandbox ? selectedMode : 'production'
  const [page, setPage] = useState(1)
  const [preview, setPreview] = useState<SignedFiche | null>(null)
  const [error, setError] = useState('')
  const [message, setMessage] = useState('')
  const [busy, setBusy] = useState(false)
  const lock = useRef(false)
  const slot = useWorkspaceToolbar()
  const list = useQuery({ queryKey: ['signed-fiches', viewer, month, year, page, mode, username],
    queryFn: ({ signal }) => signedFichesApi.list(token, month, year, page, mode, username, signal),
    retry: false, enabled: slot?.active !== false, refetchInterval: slot?.active === false ? false : 15000 })
  useDraftGuard(false, busy, 'Téléchargement des fiches signées')
  async function perform(action: () => Promise<void>) {
    if (lock.current) return
    lock.current = true; setBusy(true); setError(''); setMessage('')
    try { await action() } catch (reason) { setError(reason instanceof Error ? reason.message : 'Opération indisponible.') }
    finally { lock.current = false; setBusy(false) }
  }
  return <Paper variant="outlined" sx={{ p: { xs: 1.5, sm: 2 }, mb: 2, minWidth: 0 }}>
    <Box sx={{ display: 'flex', alignItems: 'center', flexWrap: 'wrap', gap: 1, mb: 1.5 }}>
      <Typography component="h2" variant="subtitle1" sx={{ mr: 'auto' }}>Fiches signées · {periodLabel(month, year)}</Typography>
      {localSandbox && <ToggleButtonGroup size="small" exclusive value={mode} disabled={busy} aria-label="Environnement des signatures" onChange={(_, value) => { if (value) { setMode(value); setPage(1) } }}>
        <ToggleButton value="production">Production</ToggleButton><ToggleButton value="sandbox-only">Tests SPECIMEN</ToggleButton>
      </ToggleButtonGroup>}
      {list.data?.can_export && !username && <Button size="small" startIcon={busy ? <CircularProgress size={16} /> : <DownloadOutlined />} disabled={busy || !list.data.total} onClick={() => void perform(async () => {
        downloadFile(await signedFichesApi.export(token, month, year, mode)); setMessage('ZIP du mois préparé. Téléchargement déclenché.')
      })}>Exporter les PDF signés</Button>}
      <Tooltip title="Actualiser"><span><IconButton aria-label="Actualiser les signatures" disabled={busy || list.isFetching} onClick={() => void list.refetch()}><RefreshOutlined /></IconButton></span></Tooltip>
    </Box>
    {list.data && !list.data.automatic_archive && <Alert severity="warning" sx={{ mb: 1 }}>Le retour automatique n’est pas configuré dans cet environnement. Les nouveaux envois ne seront pas récupérés automatiquement.</Alert>}
    {mode === 'sandbox-only' && <Alert severity="info" sx={{ mb: 1 }}>Documents de test SPECIMEN, séparés des archives de production.</Alert>}
    {(error || list.error) && <Alert severity="error" sx={{ mb: 1 }}>{error || list.error?.message}</Alert>}
    {message && <Alert severity="success" sx={{ mb: 1 }} onClose={() => setMessage('')}>{message}</Alert>}
    {list.isPending ? <CircularProgress size={24} /> : list.data && <>
      {!list.data.total && <Typography color="text.secondary" sx={{ py: 2 }}>Aucune demande suivie pour ce mois. Les anciens envois sans retour configuré restent à vérifier dans ClawShow.</Typography>}
      {list.data.items.map(item => <Box key={item.id} sx={{ display: 'grid', gridTemplateColumns: { xs: 'minmax(0, 1fr) auto', md: 'minmax(0, 1fr) 240px 150px auto' }, alignItems: 'center', gap: 1.5, py: 1, borderBottom: 1, borderColor: 'divider', minWidth: 0 }}>
        <Box sx={{ minWidth: 0 }}>
          <Typography variant="body2" sx={{ fontWeight: 500, overflowWrap: 'anywhere' }}>{item.employee_name}</Typography>
          <Typography variant="caption" color="text.secondary" sx={{ overflowWrap: 'anywhere' }}>{item.document_id || 'Référence en attente'}{item.file_bytes ? ` · ${Math.ceil(item.file_bytes / 1024)} Ko` : ''}</Typography>
          <Typography variant="caption" sx={{ display: { xs: 'block', md: 'none' } }}>{signedStatus[item.status] || item.status}</Typography>
          {item.last_error && <Typography variant="caption" color="warning.main" sx={{ display: 'block' }}>{item.last_error}</Typography>}
        </Box>
        <Chip size="small" variant="outlined" color={item.status === 'ready' ? 'success' : item.status === 'uncertain' ? 'warning' : 'default'} label={signedStatus[item.status] || item.status} sx={{ display: { xs: 'none', md: 'inline-flex' }, justifySelf: 'start', maxWidth: '100%' }} />
        <Typography variant="caption" color="text.secondary" sx={{ display: { xs: 'none', md: 'block' } }}>{item.completed_at ? `Retour reçu le ${new Date(`${item.completed_at}Z`).toLocaleDateString('fr-FR')}` : 'En attente'}</Typography>
        <Box sx={{ display: 'flex', gap: .25 }}>
          {item.status === 'ready' && <>
            <Tooltip title="Aperçu du PDF signé"><span><IconButton size="small" aria-label={`Aperçu du PDF signé de ${item.employee_name}`} disabled={busy} onClick={() => setPreview(item)}><VisibilityOutlined fontSize="small" /></IconButton></span></Tooltip>
            <Tooltip title="Télécharger le PDF signé"><span><IconButton size="small" aria-label={`Télécharger le PDF signé de ${item.employee_name}`} disabled={busy} onClick={() => void perform(async () => { downloadFile(await signedFichesApi.file(token, item.id)) })}><DownloadOutlined fontSize="small" /></IconButton></span></Tooltip>
          </>}
          {item.status === 'completed' && item.last_error && list.data?.can_export && <Tooltip title="Relancer la récupération du PDF"><span><IconButton size="small" aria-label="Relancer la récupération du PDF" disabled={busy} onClick={() => void perform(async () => {
            await signedFichesApi.retry(token, item.id); setMessage('Récupération relancée. Aucune nouvelle demande de signature envoyée.'); await list.refetch()
          })}><RefreshOutlined fontSize="small" /></IconButton></span></Tooltip>}
        </Box>
      </Box>)}
      {list.data.total > 20 && <Pagination sx={{ mt: 2 }} page={list.data.page} count={Math.ceil(list.data.total / 20)} disabled={busy} onChange={(_, value) => setPage(value)} />}
    </>}
    {preview && <SignedPreview key={preview.id} item={preview} token={token} onClose={() => setPreview(null)} />}
  </Paper>
}
