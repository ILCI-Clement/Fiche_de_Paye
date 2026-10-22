import { isAdmin, isSuperAdmin } from '../rolePolicy'
import { useDraftGuard, useWorkspaceGuard } from '../draftGuard'
import { useRef, useState } from 'react'
import { Alert, Box, Button, CircularProgress, IconButton, InputAdornment, Paper, Pagination, Snackbar, TextField, ToggleButton, ToggleButtonGroup, Tooltip, Typography } from '@mui/material'
import SearchRounded from '@mui/icons-material/SearchRounded'
import CloseRounded from '@mui/icons-material/CloseRounded'
import { useQuery } from '@tanstack/react-query'
import { alpha } from '@mui/material/styles'
import type { Session } from '../types'
import { downloadFile } from '../download'
import { DocumentsToolbar } from './DocumentsToolbar'
import { documentsApi, receiptPeriod } from './api'
import { DocumentUpload } from './DocumentUpload'
import { DocumentRow } from './DocumentRow'
import { DocumentPreview } from './DocumentPreview'
import type { Receipt } from './api'
import { BatchExportResult } from '../components/BatchExportResult'
import type { ExportResult } from '../components/BatchExportResult'
import { periodLabel } from './api'
import { matchesDocumentSearch } from './documentSearch'

export function DocumentsPage({ token, session }: { token: string; session: Session }) {
  const admin = isAdmin(session)
  const superAdmin = isSuperAdmin(session)
  const manager = session.role_tags.includes('Responsable')
  const [month, setMonth] = useState(() => new Date().getMonth() + 1)
  const [year, setYear] = useState(() => new Date().getFullYear())
  const [archived, setArchived] = useState(false)
  const [type, setType] = useState('all')
  const [search, setSearch] = useState('')
  const [page, setPage] = useState(1)
  const [selected, setSelected] = useState<number[]>([])
  const [preview, setPreview] = useState<Receipt | null>(null)
  const [uploadOpen, setUploadOpen] = useState(false)
  const [busy, setBusy] = useState(false)
  const guard = useWorkspaceGuard()
  useDraftGuard(false, busy, 'Documents')
  const lock = useRef(false)
  const [error, setError] = useState('')
  const [message, setMessage] = useState('')
  const [messageSeverity, setMessageSeverity] = useState<'success' | 'warning'>('success')
  const [exportResult, setExportResult] = useState<ExportResult | null>(null)
  const list = useQuery({ queryKey: ['documents', session.username], queryFn: ({ signal }) => documentsApi.list(token, signal), retry: false })
  const owners = useQuery({ queryKey: ['document-assignees', session.username], queryFn: ({ signal }) => documentsApi.assignees(token, signal), retry: false })
  function setPeriod(m: number, y: number) { guard?.request(() => { setMonth(m); setYear(y); setSelected([]); setPage(1) }) }
  async function run(action: () => Promise<void>, success: string) {
    if (lock.current) return false
    lock.current = true; setBusy(true); setError(''); setMessageSeverity('success')
    try {
      await action(); setMessage(success); setSelected([]); setPage(1)
      await list.refetch()
      return true
    } catch (e) { setError(e instanceof Error ? e.message : 'Une erreur est survenue.'); return false }
    finally { lock.current = false; setBusy(false) }
  }
  const visible = (list.data?.receipts || []).filter(r => {
    const period = receiptPeriod(r)
    return period.month === month && period.year === year && r.archived === archived
      && (type === 'all' || (type === 'unset' ? !r.document_type : r.document_type === type))
      && matchesDocumentSearch(r, search)
  })
  const pageCount = Math.max(1, Math.ceil(visible.length / 20))
  const currentPage = Math.min(page, pageCount)
  const pageRows = visible.slice((currentPage - 1) * 20, currentPage * 20)
  const selection = visible.filter(r => selected.includes(r.id))
  const totalBytes = selection.reduce((total, r) => total + r.size_bytes, 0)
  async function exportSelection() {
    if (lock.current || !selection.length) return
    const result: ExportResult = { phase: 'preparing', title: `Export Documents · ${periodLabel(month, year)}`,
      items: selection.map(receipt => `${receipt.document_number || `#${receipt.id}`} · ${receipt.original_filename}`) }
    const ids = selection.map(receipt => receipt.id)
    setExportResult(result)
    await run(async () => {
      try {
        const file = await documentsApi.export(token, ids)
        downloadFile(file)
        setExportResult({ ...result, phase: 'ready', filename: file.filename })
      } catch (error) {
        setExportResult({ ...result, phase: 'error', error: error instanceof Error ? error.message : 'Export indisponible.' })
        throw error
      }
    }, `${ids.length} document(s) préparé(s). Téléchargement déclenché.`)
  }
  return <Box className="workspace documents-workspace" sx={theme => ({
    '&.workspace': { maxWidth: 'none', px: { xs: 1.5, sm: 2 }, pt: 1.5 },
    '& .MuiTypography-h4': { fontWeight: 600 }, '& .MuiTypography-h6': { fontWeight: 600 },
    '& .MuiButton-root, & .MuiToggleButton-root, & .MuiAccordionSummary-root .MuiTypography-root': { fontWeight: 500 },
    '& .MuiTypography-body2': { color: 'text.secondary' },
    '& .MuiToggleButtonGroup-root': { gap: .75, flexWrap: 'wrap' },
    '& .MuiToggleButtonGroup-root .MuiToggleButton-root': { border: '1px solid', borderColor: alpha(theme.palette.text.primary, .12), borderRadius: '24px !important', px: 2, minHeight: 38, textTransform: 'none', color: 'text.secondary', '&.Mui-selected': { color: 'primary.main', bgcolor: alpha(theme.palette.primary.main, .12), borderColor: alpha(theme.palette.primary.main, .4) }, '&:hover': { bgcolor: alpha(theme.palette.primary.main, .08) } },
  })}>
    <DocumentsToolbar month={month} year={year} busy={busy} canUpload={!!owners.data && !owners.error} refreshing={list.isFetching} onPeriod={setPeriod}
      onUpload={() => { setError(''); setUploadOpen(true) }} onRefresh={() => void list.refetch()} />
    {owners.error ? <Alert severity="error" sx={{ mb: 2 }}>{owners.error.message}<Button onClick={() => void owners.refetch()}>Réessayer</Button></Alert>
      : owners.isPending ? <CircularProgress size={22} /> : null}
    {uploadOpen && owners.data && <DocumentUpload token={token} username={session.username} currentDisplayName={session.name} admin={admin} accounts={owners.data.accounts || owners.data.users.map(username => ({ username }))} month={month} year={year} busy={busy} run={run} requestError={error} onClose={() => setUploadOpen(false)} />}
    <Paper sx={{ p: 1.5, mb: 1.5, borderRadius: 2 }}>
      <Box sx={{ display: 'flex', alignItems: 'center', flexWrap: 'wrap', gap: 1 }}>
        <TextField placeholder="Rechercher un document · nom, numéro, employé…"
          value={search} disabled={busy} onChange={event => { setSearch(event.target.value); setSelected([]); setPage(1) }}
          size="small" sx={{ flex: '1 1 280px', minWidth: 0, '& .MuiOutlinedInput-root': { borderRadius: 2.5 } }}
          slotProps={{ htmlInput: { 'aria-label': 'Rechercher un document dans le mois et le dossier sélectionnés' }, input: {
            startAdornment: <InputAdornment position="start"><SearchRounded fontSize="small" /></InputAdornment>,
            endAdornment: search ? <InputAdornment position="end"><Tooltip title="Effacer la recherche"><IconButton
              aria-label="Effacer la recherche" size="small" disabled={busy} onClick={() => { setSearch(''); setSelected([]); setPage(1) }}><CloseRounded fontSize="small" /></IconButton></Tooltip></InputAdornment> : undefined,
          } }} />
      </Box>
      <Box sx={{ display: 'flex', alignItems: 'center', flexWrap: 'wrap', gap: 1, my: 1 }}>
        <ToggleButtonGroup exclusive value={archived ? 'archives' : 'active'} size="small" disabled={busy} aria-label="Dossier" onChange={(_, value) => { if (value) { setArchived(value === 'archives'); setSelected([]); setPage(1) } }}>
          <ToggleButton value="active">À traiter</ToggleButton><ToggleButton value="archives">Archives</ToggleButton>
        </ToggleButtonGroup>
        <ToggleButtonGroup exclusive value={type} size="small" disabled={busy} aria-label="Filtrer par type" sx={{ flexWrap: 'wrap' }} onChange={(_, value) => { if (value) { setType(value); setSelected([]); setPage(1) } }}>
          {[['all', 'Tous'], ['transport', 'Transport'], ['autre', 'Autre'], ['unset', 'Non renseigné']].map(([value, label]) => <ToggleButton key={value} value={value}>{label}</ToggleButton>)}
        </ToggleButtonGroup>
        <Typography variant="caption" color="text.secondary" aria-live="polite" sx={{ ml: { sm: 'auto' } }}>{visible.length} document(s)</Typography>
      </Box>
      <Box sx={{ display: 'flex', flexWrap: 'wrap', gap: 1, alignItems: 'center', '& .MuiButton-root': { minHeight: 32, py: .5 } }}>
        <Button disabled={busy || !visible.length || !!list.error} onClick={() => setSelected(visible.map(r => r.id))}>Tout sélectionner ({visible.length})</Button>
        <Button disabled={busy || !selection.length} onClick={() => setSelected([])}>Tout désélectionner</Button>
        <Button variant="contained" disabled={busy || !selection.length || selection.length > 100 || totalBytes > 100 * 1024 * 1024 || !!list.error} onClick={() => void exportSelection()}>Télécharger le ZIP ({selection.length})</Button>
        <Typography variant="caption" color="text.secondary">100 fichiers / 100 Mo maximum · {(totalBytes / 1024 / 1024).toFixed(1)} Mo sélectionnés</Typography>
      </Box>
    </Paper>
    {exportResult && <BatchExportResult result={exportResult} />}
    {error && <Alert severity="error" sx={{ mb: 2 }} onClose={() => setError('')}>{error}</Alert>}
    {list.error ? <Alert severity="error">{list.error.message}<Button onClick={() => void list.refetch()}>Réessayer</Button></Alert> : list.isPending ? <CircularProgress size={28} /> :
      <Box sx={{ display: 'flex', flexDirection: 'column', gap: .75 }}>
        {!visible.length && <Typography color="text.secondary">{search.trim() ? 'Aucun document ne correspond à la recherche et aux filtres.' : 'Aucun document dans ce dossier pour ce mois.'}</Typography>}
        {pageRows.map(r => <DocumentRow key={r.id} receipt={r} admin={admin} canArchive={admin || (!manager && (r.uploaded_by === session.username || r.employee_username === session.username))} canDelete={superAdmin && r.archived} selected={selected.includes(r.id)} busy={busy}
          onPreview={() => setPreview(r)}
          onSelect={() => setSelected(previous => previous.includes(r.id) ? previous.filter(id => id !== r.id) : [...previous, r.id])}
          onDownload={() => void run(async () => downloadFile(await documentsApi.file(token, r.id)), 'Document prêt à télécharger.')}
          onArchive={() => run(async () => { await documentsApi.archive(token, r.id, !r.archived) }, r.archived ? 'Document restauré.' : 'Document archivé.')}
          onDelete={async confirmation => {
            const result = await documentsApi.remove(token, r.id, confirmation)
            setMessageSeverity(result.cleanup_pending ? 'warning' : 'success')
            setMessage(result.cleanup_pending ? 'Document retiré du registre ; le nettoyage du fichier reste à terminer.' : 'Document supprimé.')
            setSelected(previous => previous.filter(id => id !== r.id)); await list.refetch()
          }}
          onPeriod={(m, y) => run(async () => { await documentsApi.period(token, r.id, m, y) }, 'Mois de rattachement mis à jour.')} />)}
      </Box>}
    {!list.isPending && !list.error && visible.length > 0 && <Box sx={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', flexWrap: 'wrap', gap: 1, py: 1.5 }}>
      <Typography variant="caption" color="text.secondary">{(currentPage - 1) * 20 + 1}–{Math.min(currentPage * 20, visible.length)} sur {visible.length} · 20 par page</Typography>
      <Pagination count={pageCount} page={currentPage} disabled={busy} onChange={(_, value) => setPage(value)} color="primary" size="small" aria-label="Pagination des documents" />
    </Box>}
    {preview && <DocumentPreview key={preview.id} token={token} receipt={preview} onClose={() => setPreview(null)} />}
    <Snackbar open={!!message} autoHideDuration={5000} onClose={() => setMessage('')}><Alert severity={messageSeverity} variant="filled" onClose={() => setMessage('')}>{message}</Alert></Snackbar>
  </Box>
}
