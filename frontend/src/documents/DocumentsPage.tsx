import { useDraftGuard, useWorkspaceGuard } from '../draftGuard'
import { useRef, useState } from 'react'
import { Alert, Box, Button, CircularProgress, IconButton, InputAdornment, Paper, Snackbar, TextField, ToggleButton, ToggleButtonGroup, Tooltip, Typography } from '@mui/material'
import SearchRounded from '@mui/icons-material/SearchRounded'
import CloseRounded from '@mui/icons-material/CloseRounded'
import ZoomOutRounded from '@mui/icons-material/ZoomOutRounded'
import ZoomInRounded from '@mui/icons-material/ZoomInRounded'
import UploadFileRounded from '@mui/icons-material/UploadFileRounded'
import { useQuery } from '@tanstack/react-query'
import { alpha } from '@mui/material/styles'
import type { Session } from '../types'
import { downloadFile } from '../download'
import { PeriodSelector } from '../attendance/PeriodSelector'
import { documentsApi, receiptPeriod } from './api'
import { DocumentUpload } from './DocumentUpload'
import { DocumentCard } from './DocumentCard'
import type { DocumentDensity } from './DocumentCard'
import { DocumentPreview } from './DocumentPreview'
import type { Receipt } from './api'
import { BatchExportResult } from '../components/BatchExportResult'
import type { ExportResult } from '../components/BatchExportResult'
import { periodLabel } from './api'
import { matchesDocumentSearch } from './documentSearch'

const densities: DocumentDensity[] = ['compact', 'standard', 'comfortable']
const densityLabels = { compact: 'Compact', standard: 'Standard', comfortable: 'Confort' }
const densityKey = 'presence-document-density'
function initialDensity(): DocumentDensity {
  try {
    const saved = localStorage.getItem(densityKey)
    return densities.find(value => value === saved) || 'standard'
  } catch { return 'standard' }
}

export function DocumentsPage({ token, session }: { token: string; session: Session }) {
  const admin = session.role_tags.includes('Admin')
  const manager = session.role_tags.includes('Responsable')
  const [month, setMonth] = useState(() => new Date().getMonth() + 1)
  const [year, setYear] = useState(() => new Date().getFullYear())
  const [archived, setArchived] = useState(false)
  const [type, setType] = useState('all')
  const [search, setSearch] = useState('')
  const [density, setDensity] = useState<DocumentDensity>(initialDensity)
  function changeDensity(direction: number) {
    const next = densities[Math.max(0, Math.min(densities.length - 1, densities.indexOf(density) + direction))]
    setDensity(next)
    try { localStorage.setItem(densityKey, next) } catch { /* Display preferences are optional. */ }
  }
  const [selected, setSelected] = useState<number[]>([])
  const [preview, setPreview] = useState<Receipt | null>(null)
  const [uploadOpen, setUploadOpen] = useState(false)
  const [busy, setBusy] = useState(false)
  const guard = useWorkspaceGuard()
  useDraftGuard(false, busy, 'Documents')
  const lock = useRef(false)
  const [error, setError] = useState('')
  const [message, setMessage] = useState('')
  const [exportResult, setExportResult] = useState<ExportResult | null>(null)
  const list = useQuery({ queryKey: ['documents', session.username], queryFn: ({ signal }) => documentsApi.list(token, signal), retry: false })
  const owners = useQuery({ queryKey: ['document-assignees', session.username], queryFn: ({ signal }) => documentsApi.assignees(token, signal), retry: false })
  function setPeriod(m: number, y: number) { guard?.request(() => { setMonth(m); setYear(y); setSelected([]) }) }
  async function run(action: () => Promise<void>, success: string) {
    if (lock.current) return false
    lock.current = true; setBusy(true); setError('')
    try {
      await action(); setMessage(success); setSelected([])
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
    '& .MuiTypography-h4': { fontWeight: 600 }, '& .MuiTypography-h6': { fontWeight: 600 },
    '& .MuiButton-root, & .MuiToggleButton-root, & .MuiAccordionSummary-root .MuiTypography-root': { fontWeight: 500 },
    '& .MuiTypography-body2': { color: 'text.secondary' },
    '& .MuiToggleButtonGroup-root': { gap: .75, flexWrap: 'wrap' },
    '& .MuiToggleButtonGroup-root .MuiToggleButton-root': { border: '1px solid', borderColor: alpha(theme.palette.text.primary, .12), borderRadius: '24px !important', px: 2, minHeight: 38, textTransform: 'none', color: 'text.secondary', '&.Mui-selected': { color: 'primary.main', bgcolor: alpha(theme.palette.primary.main, .12), borderColor: alpha(theme.palette.primary.main, .4) }, '&:hover': { bgcolor: alpha(theme.palette.primary.main, .08) } },
  })}>
    <Box className="page-heading"><Box><Typography variant="overline" color="text.secondary" sx={{ fontWeight: 500, letterSpacing: 2 }}>DOCUMENTS</Typography>
      <Typography variant="h4" component="h1">Documents et justificatifs</Typography>
      <Typography color="text.secondary" sx={{ mt: 1 }}>Déposez et retrouvez vos documents par mois.</Typography>
    </Box></Box>
    <Paper component="section" aria-label="Mois des documents" sx={{ p: 2, mb: 2, position: 'sticky', top: 0, zIndex: 1, display: 'flex', flexWrap: 'wrap', alignItems: 'center', justifyContent: 'space-between', gap: 1.5, border: 1, borderColor: 'divider', borderRadius: 3 }}>
      <PeriodSelector month={month} year={year} disabled={busy} onChange={setPeriod} />
      <Button variant="contained" startIcon={<UploadFileRounded />} disabled={busy || !owners.data || !!owners.error}
        onClick={() => { setError(''); setUploadOpen(true) }}>Déposer un document</Button>
    </Paper>
    {owners.error ? <Alert severity="error" sx={{ mb: 2 }}>{owners.error.message}<Button onClick={() => void owners.refetch()}>Réessayer</Button></Alert>
      : owners.isPending ? <CircularProgress size={22} /> : null}
    {uploadOpen && owners.data && <DocumentUpload token={token} username={session.username} admin={admin} users={owners.data.users} month={month} year={year} busy={busy} run={run} requestError={error} onClose={() => setUploadOpen(false)} />}
    <Paper sx={{ p: { xs: 2, sm: 2.5 }, my: 2, borderRadius: 3 }}>
      <Box sx={{ display: 'flex', alignItems: 'center', flexWrap: 'wrap', gap: 2 }}>
        <Box sx={{ flex: 1 }}><Typography variant="h6">{archived ? 'Archives' : 'Tous les justificatifs'}</Typography><Typography variant="body2" color="text.secondary">{visible.length} document(s) correspondant aux filtres</Typography></Box>
        <Box role="group" aria-label="Taille des cartes" sx={{ display: 'flex', alignItems: 'center', gap: .5 }}>
          <Tooltip title="Réduire les cartes"><span><IconButton aria-label="Réduire les cartes" disabled={density === 'compact'} onClick={() => changeDensity(-1)}><ZoomOutRounded /></IconButton></span></Tooltip>
          <Typography variant="body2" aria-live="polite" sx={{ minWidth: 64, textAlign: 'center' }}>{densityLabels[density]}</Typography>
          <Tooltip title="Agrandir les cartes"><span><IconButton aria-label="Agrandir les cartes" disabled={density === 'comfortable'} onClick={() => changeDensity(1)}><ZoomInRounded /></IconButton></span></Tooltip>
        </Box>
        <Button disabled={busy || list.isFetching} onClick={() => void list.refetch()}>Actualiser</Button>
      </Box>
      <TextField placeholder="Rechercher un document · nom, numéro, employé…"
        value={search} disabled={busy} onChange={event => { setSearch(event.target.value); setSelected([]) }}
        fullWidth size="small" sx={{ mt: 2, '& .MuiOutlinedInput-root': { borderRadius: 2.5 } }}
        slotProps={{ htmlInput: { 'aria-label': 'Rechercher un document dans le mois et le dossier sélectionnés' }, input: {
          startAdornment: <InputAdornment position="start"><SearchRounded fontSize="small" /></InputAdornment>,
          endAdornment: search ? <InputAdornment position="end"><Tooltip title="Effacer la recherche"><IconButton
            aria-label="Effacer la recherche" size="small" disabled={busy} onClick={() => { setSearch(''); setSelected([]) }}><CloseRounded fontSize="small" /></IconButton></Tooltip></InputAdornment> : undefined,
        } }} />
      <Box sx={{ display: 'flex', flexWrap: 'wrap', gap: 2, my: 2 }}>
        <ToggleButtonGroup exclusive value={archived ? 'archives' : 'active'} size="small" disabled={busy} aria-label="Dossier" onChange={(_, value) => { if (value) { setArchived(value === 'archives'); setSelected([]) } }}>
          <ToggleButton value="active">À traiter</ToggleButton><ToggleButton value="archives">Archives</ToggleButton>
        </ToggleButtonGroup>
        <ToggleButtonGroup exclusive value={type} size="small" disabled={busy} aria-label="Filtrer par type" sx={{ flexWrap: 'wrap' }} onChange={(_, value) => { if (value) { setType(value); setSelected([]) } }}>
          {[['all', 'Tous'], ['transport', 'Transport'], ['autre', 'Autre'], ['unset', 'Non renseigné']].map(([value, label]) => <ToggleButton key={value} value={value}>{label}</ToggleButton>)}
        </ToggleButtonGroup>
      </Box>
      <Box sx={{ display: 'flex', flexWrap: 'wrap', gap: 1, alignItems: 'center' }}>
        <Button disabled={busy || !visible.length || !!list.error} onClick={() => setSelected(visible.map(r => r.id))}>Tout sélectionner</Button>
        <Button disabled={busy || !selection.length} onClick={() => setSelected([])}>Tout désélectionner</Button>
        <Button variant="contained" disabled={busy || !selection.length || selection.length > 100 || totalBytes > 100 * 1024 * 1024 || !!list.error} onClick={() => void exportSelection()}>Télécharger le ZIP ({selection.length})</Button>
        <Typography variant="caption" color="text.secondary">100 fichiers / 100 Mo maximum · {(totalBytes / 1024 / 1024).toFixed(1)} Mo sélectionnés</Typography>
      </Box>
    </Paper>
    {exportResult && <BatchExportResult result={exportResult} />}
    {error && <Alert severity="error" sx={{ mb: 2 }} onClose={() => setError('')}>{error}</Alert>}
    {list.error ? <Alert severity="error">{list.error.message}<Button onClick={() => void list.refetch()}>Réessayer</Button></Alert> : list.isPending ? <CircularProgress size={28} /> :
      <Box sx={{ display: 'grid', gridTemplateColumns: density === 'compact'
        ? { xs: '1fr', sm: 'repeat(2, minmax(0, 1fr))', lg: 'repeat(3, minmax(0, 1fr))', xl: 'repeat(4, minmax(0, 1fr))' }
        : density === 'comfortable' ? { xs: '1fr', lg: 'repeat(2, minmax(0, 1fr))' }
          : { xs: '1fr', md: 'repeat(2, minmax(0, 1fr))', xl: 'repeat(3, minmax(0, 1fr))' }, gap: density === 'compact' ? 1 : density === 'comfortable' ? 3 : 2 }}>
        {!visible.length && <Typography color="text.secondary">{search.trim() ? 'Aucun document ne correspond à la recherche et aux filtres.' : 'Aucun document dans ce dossier pour ce mois.'}</Typography>}
        {visible.map(r => <DocumentCard key={r.id} receipt={r} density={density} admin={admin} canDelete={admin || (!manager && r.uploaded_by === session.username)} selected={selected.includes(r.id)} busy={busy}
          onPreview={() => setPreview(r)}
          onSelect={() => setSelected(previous => previous.includes(r.id) ? previous.filter(id => id !== r.id) : [...previous, r.id])}
          onDownload={() => void run(async () => downloadFile(await documentsApi.file(token, r.id)), 'Document prêt à télécharger.')}
          onArchive={() => void run(async () => { await documentsApi.archive(token, r.id) }, 'Document déplacé dans les archives.')}
          onDelete={() => run(async () => { await documentsApi.remove(token, r.id) }, 'Document supprimé.')}
          onPeriod={(m, y) => run(async () => { await documentsApi.period(token, r.id, m, y) }, 'Mois de rattachement mis à jour.')} />)}
      </Box>}
    {preview && <DocumentPreview key={preview.id} token={token} receipt={preview} onClose={() => setPreview(null)} />}
    <Snackbar open={!!message} autoHideDuration={5000} onClose={() => setMessage('')}><Alert severity="success" variant="filled" onClose={() => setMessage('')}>{message}</Alert></Snackbar>
  </Box>
}
