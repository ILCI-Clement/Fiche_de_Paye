import { Fragment, useCallback, useRef, useState } from 'react'
import { Alert, Box, Button, Checkbox, Chip, CircularProgress, IconButton, Pagination, Paper, Tab, Tabs, ToggleButton, ToggleButtonGroup, Tooltip, Typography } from '@mui/material'
import DownloadOutlined from '@mui/icons-material/DownloadOutlined'
import VisibilityOutlined from '@mui/icons-material/VisibilityOutlined'
import RefreshOutlined from '@mui/icons-material/RefreshOutlined'
import ExpandMoreRounded from '@mui/icons-material/ExpandMoreRounded'
import ChevronRightRounded from '@mui/icons-material/ChevronRightRounded'
import DeleteOutlineRounded from '@mui/icons-material/DeleteOutlineRounded'
import { useQuery, useQueryClient } from '@tanstack/react-query'
import { signedFichesApi, signedStatus } from './signedFichesApi'
import type { SignedFiche } from './signedFichesApi'
import { DocumentPreview } from '../documents/DocumentPreview'
import { downloadFile } from '../download'
import { periodLabel } from '../documents/api'
import { useDraftGuard } from '../draftGuard'
import { useWorkspaceToolbar } from '../workspaceToolbar'
import { useCapabilities } from '../capabilities'
import { FileBatchDialog } from '../components/FileBatchDialog'
import { fileActionLabels } from '../fileBatch'
import type { FileAction } from '../fileBatch'
import ArchiveOutlined from '@mui/icons-material/ArchiveOutlined'
import UnarchiveOutlined from '@mui/icons-material/UnarchiveOutlined'
import { PermanentDeleteDialog } from '../components/PermanentDeleteDialog'

function SignedPreview({ item, token, onClose }: { item: SignedFiche; token: string; onClose: () => void }) {
  const loadFile = useCallback((signal: AbortSignal) => signedFichesApi.file(token, item.id, signal), [token, item.id])
  return <DocumentPreview token={token} receipt={{ id: 0, original_filename: `${item.employee_name} · ${periodLabel(item.fiche_month, item.fiche_year)} · signé.pdf` }} loadFile={loadFile} onClose={onClose} />
}

export function SignedFichesPanel({ token, month, year, username, viewer }: { token: string; month?: number; year: number; username?: string; viewer: string }) {
  const [archived, setArchived] = useState(false)
  const [selectedFiles, setSelectedFiles] = useState<{ scope: string; ids: string[] }>({ scope: '', ids: [] })
  const [batch, setBatch] = useState<{ action: FileAction; items: SignedFiche[]; archived: boolean } | null>(null)
  const [selectedMode, setMode] = useState('production')
  const localSandbox = useCapabilities().data?.environment === 'local-snapshot'
  const mode = localSandbox ? selectedMode : 'production'
  const [page, setPage] = useState(1)
  const [collapsedMonths, setCollapsedMonths] = useState<Record<number, boolean>>({})
  const [preview, setPreview] = useState<SignedFiche | null>(null)
  const [deleting, setDeleting] = useState<SignedFiche | null>(null)
  const queryClient = useQueryClient()
  const [error, setError] = useState('')
  const [message, setMessage] = useState('')
  const [busy, setBusy] = useState(false)
  const lock = useRef(false)
  const slot = useWorkspaceToolbar()
  const list = useQuery({ queryKey: ['signed-fiches', viewer, month, year, page, mode, username, archived],
    queryFn: ({ signal }) => signedFichesApi.list(token, month, year, page, mode, username, signal, archived),
    retry: false, enabled: slot?.active !== false, refetchInterval: slot?.active === false ? false : 15000 })
  const selectionScope = JSON.stringify([viewer, year, month, mode, username, page, archived])
  const eligible = list.data?.items.filter(item => item.status === 'ready') || []
  const selection = selectedFiles.scope === selectionScope ? eligible.filter(item => selectedFiles.ids.includes(item.id)) : []
  const allSelected = eligible.length > 0 && selection.length === eligible.length
  const canManage = !!list.data?.can_export && !username
  const toggleFile = (id: string) => setSelectedFiles({ scope: selectionScope,
    ids: selection.some(item => item.id === id) ? selection.filter(item => item.id !== id).map(item => item.id) : [...selection.map(item => item.id), id] })
  useDraftGuard(false, busy, 'Opération sur les fiches signées')
  async function perform(action: () => Promise<void>) {
    if (lock.current) return
    lock.current = true; setBusy(true); setError(''); setMessage('')
    try { await action() } catch (reason) { setError(reason instanceof Error ? reason.message : 'Opération indisponible.') }
    finally { lock.current = false; setBusy(false) }
  }
  return <Paper variant="outlined" sx={{ p: { xs: 1.5, sm: 2 }, mb: 2, minWidth: 0 }}>
    <Box sx={{ display: 'flex', alignItems: 'center', flexWrap: 'wrap', gap: 1, mb: 1.5 }}>
      <Typography component="h2" variant="subtitle1" sx={{ mr: 'auto' }}>Fiches signées · {month ? periodLabel(month, year) : year}</Typography>
      {localSandbox && <ToggleButtonGroup size="small" exclusive value={mode} disabled={busy} aria-label="Environnement des signatures" onChange={(_, value) => { if (value) { setMode(value); setPage(1); setSelectedFiles({ scope: '', ids: [] }) } }}>
        <ToggleButton value="production">Production</ToggleButton><ToggleButton value="sandbox-only">Tests SPECIMEN</ToggleButton>
      </ToggleButtonGroup>}
      {list.data?.can_export && !username && month && <Button size="small" startIcon={busy ? <CircularProgress size={16} /> : <DownloadOutlined />} disabled={busy || !list.data.total} onClick={() => void perform(async () => {
        downloadFile(await signedFichesApi.export(token, month, year, mode)); setMessage('ZIP du mois préparé. Téléchargement déclenché.')
      })}>Exporter les PDF signés</Button>}
      {list.data?.can_export && !username && !month && <Button size="small" variant="outlined"
        startIcon={busy ? <CircularProgress size={16} /> : <DownloadOutlined />} disabled={busy || !list.data.total}
        onClick={() => void perform(async () => {
          downloadFile(await signedFichesApi.exportYear(token, year, mode)); setMessage('ZIP annuel préparé, classé par mois.')
        })}>Télécharger l’année</Button>}
      <Tooltip title="Actualiser"><span><IconButton aria-label="Actualiser les signatures" disabled={busy || list.isFetching} onClick={() => void list.refetch()}><RefreshOutlined /></IconButton></span></Tooltip>
    </Box>
    <Box sx={{ display: 'flex', gap: 1, alignItems: 'center', flexWrap: 'wrap', mb: 1 }}>
      <Tabs value={archived ? 'archives' : 'active'} onChange={(_, value) => { setArchived(value === 'archives'); setPage(1); setSelectedFiles({ scope: '', ids: [] }) }}
        aria-label="Dossier des PDF signés" sx={{ minHeight: 36, '& .MuiTab-root': { minHeight: 36 } }}>
        <Tab value="active" label="Normal" disabled={busy} /><Tab value="archives" label="Archives" disabled={busy} />
      </Tabs>
      {canManage && <>
        <Button size="small" disabled={busy || !eligible.length || !!list.error}
          onClick={() => setSelectedFiles({ scope: selectionScope, ids: allSelected ? [] : eligible.map(item => item.id) })}>
          {allSelected ? 'Tout désélectionner' : 'Sélectionner la page'}</Button>
        {selection.length > 0 && <>
          <Button size="small" disabled={busy || !!list.error} onClick={() => setBatch({ action: archived ? 'restore' : 'archive', items: [...selection], archived })}>
            {archived ? 'Restaurer' : 'Archiver'} ({selection.length})</Button>
          {list.data?.can_delete && <Button size="small" color="error" disabled={busy || !!list.error}
            onClick={() => setBatch({ action: 'delete', items: [...selection], archived })}>Supprimer ({selection.length})</Button>}
        </>}
      </>}
    </Box>
    {list.data && !list.data.automatic_archive && <Alert severity="warning" sx={{ mb: 1 }}>Le retour automatique n’est pas configuré dans cet environnement. Les nouveaux envois ne seront pas récupérés automatiquement.</Alert>}
    {mode === 'sandbox-only' && <Alert severity="info" sx={{ mb: 1 }}>Documents de test SPECIMEN, séparés des archives de production.</Alert>}
    {(error || list.error) && <Alert severity="error" sx={{ mb: 1 }}>{error || list.error?.message}</Alert>}
    {message && <Alert severity="success" sx={{ mb: 1 }} onClose={() => setMessage('')}>{message}</Alert>}
    {list.isPending ? <CircularProgress size={24} /> : list.data && <>
      {!list.data.total && <Typography color="text.secondary" sx={{ py: 2 }}>{archived ? 'Aucun PDF dans Archives pour cette période.' : `Aucune demande suivie pour ${month ? 'ce mois' : 'cette année'}. Les anciens envois sans retour configuré restent à vérifier dans ClawShow.`}</Typography>}
      {list.data.items.map((item, index) => <Fragment key={item.id}>
        {!month && (index === 0 || list.data.items[index - 1].fiche_month !== item.fiche_month) && <Box sx={{ display: 'flex', alignItems: 'center', gap: 1, mt: 2, mb: .5, px: 1, py: .75, bgcolor: 'action.hover', borderRadius: 1 }}>
          <Typography component="h3" variant="subtitle2" sx={{ mr: 'auto' }}>
            <Button color="inherit" size="small" aria-expanded={!collapsedMonths[item.fiche_month]} aria-label={`${collapsedMonths[item.fiche_month] ? 'Déplier' : 'Replier'} ${periodLabel(item.fiche_month, year)}`} startIcon={collapsedMonths[item.fiche_month] ? <ChevronRightRounded /> : <ExpandMoreRounded />} onClick={() => setCollapsedMonths(current => ({ ...current, [item.fiche_month]: !current[item.fiche_month] }))} sx={{ textTransform: 'capitalize', fontWeight: 600 }}>{periodLabel(item.fiche_month, year)}</Button>
          </Typography>
          {list.data.can_export && !username && <Button size="small" startIcon={<DownloadOutlined />} disabled={busy} onClick={() => void perform(async () => { downloadFile(await signedFichesApi.export(token, item.fiche_month, year, mode)) })}>Exporter le mois</Button>}
        </Box>}
        {(month || !collapsedMonths[item.fiche_month]) && <Box sx={{ display: 'grid', gridTemplateColumns: { xs: 'minmax(0, 1fr) auto', md: 'minmax(0, 1fr) 240px 150px auto' }, alignItems: 'center', gap: 1.5, py: 1, borderBottom: 1, borderColor: 'divider', minWidth: 0 }}>
        <Box sx={{ minWidth: 0, display: 'flex', alignItems: 'center', gap: .5 }}>
          {canManage && <Checkbox size="small" checked={selection.some(file => file.id === item.id)} disabled={busy || item.status !== 'ready'}
            onChange={() => toggleFile(item.id)} slotProps={{ input: { 'aria-label': `Sélectionner le PDF signé de ${item.employee_name} (${item.document_id || item.id})` } }} />}
          <Box sx={{ minWidth: 0 }}><Typography variant="body2" sx={{ fontWeight: 500, overflowWrap: 'anywhere' }}>{item.employee_name}</Typography>
          <Typography variant="caption" color="text.secondary" sx={{ overflowWrap: 'anywhere' }}>{item.document_id || 'Référence en attente'}{item.file_bytes ? ` · ${Math.ceil(item.file_bytes / 1024)} Ko` : ''}</Typography>
          <Typography variant="caption" sx={{ display: { xs: 'block', md: 'none' } }}>{signedStatus[item.status] || item.status}</Typography>
          {item.last_error && <Typography variant="caption" color="warning.main" sx={{ display: 'block' }}>{item.last_error}</Typography>}
        </Box>
        </Box>
        <Chip size="small" variant="outlined" color={item.status === 'ready' ? 'success' : item.status === 'uncertain' ? 'warning' : 'default'} label={signedStatus[item.status] || item.status} sx={{ display: { xs: 'none', md: 'inline-flex' }, justifySelf: 'start', maxWidth: '100%' }} />
        <Typography variant="caption" color="text.secondary" sx={{ display: { xs: 'none', md: 'block' } }}>{item.completed_at ? `Retour reçu le ${new Date(`${item.completed_at}Z`).toLocaleDateString('fr-FR')}` : 'En attente'}</Typography>
        <Box sx={{ display: 'flex', gap: .25 }}>
          {item.status === 'ready' && <>
            <Tooltip title="Aperçu du PDF signé"><span><IconButton size="small" aria-label={`Aperçu du PDF signé de ${item.employee_name}`} disabled={busy} onClick={() => setPreview(item)}><VisibilityOutlined fontSize="small" /></IconButton></span></Tooltip>
            <Tooltip title="Télécharger le PDF signé"><span><IconButton size="small" aria-label={`Télécharger le PDF signé de ${item.employee_name}`} disabled={busy} onClick={() => void perform(async () => { downloadFile(await signedFichesApi.file(token, item.id)) })}><DownloadOutlined fontSize="small" /></IconButton></span></Tooltip>
          </>}
          {item.status === 'ready' && canManage && <Tooltip title={archived ? 'Restaurer' : 'Archiver'}><span><IconButton size="small"
            aria-label={`${archived ? 'Restaurer' : 'Archiver'} le PDF signé de ${item.employee_name}`} disabled={busy}
            onClick={() => setBatch({ action: archived ? 'restore' : 'archive', items: [item], archived })}>
            {archived ? <UnarchiveOutlined fontSize="small" /> : <ArchiveOutlined fontSize="small" />}
          </IconButton></span></Tooltip>}
          {item.status === 'ready' && list.data.can_delete && <Tooltip title="Supprimer définitivement"><span><IconButton size="small" color="error" aria-label={`Supprimer le PDF signé de ${item.employee_name}`} disabled={busy} onClick={() => setDeleting(item)}><DeleteOutlineRounded fontSize="small" /></IconButton></span></Tooltip>}
          {item.status === 'completed' && item.last_error && list.data?.can_export && <Tooltip title="Relancer la récupération du PDF"><span><IconButton size="small" aria-label="Relancer la récupération du PDF" disabled={busy} onClick={() => void perform(async () => {
            await signedFichesApi.retry(token, item.id); setMessage('Récupération relancée. Aucune nouvelle demande de signature envoyée.'); await list.refetch()
          })}><RefreshOutlined fontSize="small" /></IconButton></span></Tooltip>}
        </Box>
      </Box>}</Fragment>)}
      {list.data.total > 20 && <Pagination sx={{ mt: 2 }} page={list.data.page} count={Math.ceil(list.data.total / 20)} disabled={busy} onChange={(_, value) => { setPage(value); setSelectedFiles({ scope: '', ids: [] }) }} />}
    </>}
    {batch && <FileBatchDialog action={batch.action} items={batch.items.map(item => `${item.employee_name} · ${periodLabel(item.fiche_month, item.fiche_year)} · ${item.document_id || item.id}`)}
      onClose={() => setBatch(null)} onConfirm={async confirmation => {
        if (lock.current) throw new Error('Une opération est déjà en cours.')
        lock.current = true; setBusy(true); setError(''); setMessage('')
        try {
          const result = await signedFichesApi.batch(token, batch.items.map(item => item.id), batch.action, batch.archived, confirmation)
          setSelectedFiles({ scope: '', ids: [] }); setPage(1)
          if (result.cleanup_pending) setError('Fichiers retirés. Nettoyage à terminer par la maintenance.')
          else setMessage(`${result.count} fichier(s) : ${fileActionLabels[batch.action].toLowerCase()}.`)
          await Promise.all(['signed-fiches', 'person-profile', 'work-center'].map(key => queryClient.invalidateQueries({ queryKey: [key] })))
        } catch (failure) { await list.refetch(); throw failure }
        finally { lock.current = false; setBusy(false) }
      }} />}
    {preview && <SignedPreview key={preview.id} item={preview} token={token} onClose={() => setPreview(null)} />}
    {deleting && <PermanentDeleteDialog title="Supprimer définitivement ce PDF signé ?" expected="SUPPRIMER"
      description={`${deleting.employee_name} · ${periodLabel(deleting.fiche_month, deleting.fiche_year)}. Le PDF archivé et sa ligne seront supprimés de cette application. L’original ClawShow reste inchangé.`}
      onClose={() => setDeleting(null)} onDeleted={() => setDeleting(null)} onConfirm={async confirmation => {
        setBusy(true); setError(''); setMessage('')
        try {
          const result = await signedFichesApi.delete(token, deleting.id, confirmation)
          setPreview(current => current?.id === deleting.id ? null : current)
          if (result.cleanup_pending) setError('Document retiré. Le nettoyage du fichier doit être terminé par la maintenance.')
          else setMessage('PDF signé supprimé.')
          await Promise.all(['signed-fiches', 'person-profile', 'work-center'].map(key => queryClient.invalidateQueries({ queryKey: [key] })))
        } finally { setBusy(false) }
      }} />}
  </Paper>
}
