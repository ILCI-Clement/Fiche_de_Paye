import { useDraftGuard, useWorkspaceGuard } from '../draftGuard'
import { useRef, useState } from 'react'
import { Alert, Box, Button, CircularProgress, Paper, Snackbar, ToggleButton, ToggleButtonGroup, Typography } from '@mui/material'
import { useQuery } from '@tanstack/react-query'
import type { Session } from '../types'
import { downloadFile } from '../download'
import { PeriodSelector } from '../attendance/PeriodSelector'
import { documentsApi, receiptPeriod } from './api'
import { DocumentUpload } from './DocumentUpload'
import { DocumentCard } from './DocumentCard'

export function DocumentsPage({ token, session }: { token: string; session: Session }) {
  const admin = session.role_tags.includes('Admin')
  const manager = session.role_tags.includes('Responsable')
  const [month, setMonth] = useState(() => new Date().getMonth() + 1)
  const [year, setYear] = useState(() => new Date().getFullYear())
  const [archived, setArchived] = useState(false)
  const [type, setType] = useState('all')
  const [selected, setSelected] = useState<number[]>([])
  const [busy, setBusy] = useState(false)
  const guard = useWorkspaceGuard()
  useDraftGuard(false, busy, 'Documents')
  const lock = useRef(false)
  const [error, setError] = useState('')
  const [message, setMessage] = useState('')
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
  })
  const selection = visible.filter(r => selected.includes(r.id))
  const totalBytes = selection.reduce((total, r) => total + r.size_bytes, 0)
  return <Box className="workspace">
    <Box className="page-heading"><Box><Typography variant="overline" color="primary.main" sx={{ fontWeight: 700 }}>DOCUMENTS</Typography>
      <Typography variant="h4" component="h1">Documents et justificatifs</Typography>
      <Typography color="text.secondary" sx={{ mt: 1 }}>Déposez et retrouvez vos documents par mois.</Typography>
    </Box></Box>
    <Paper component="section" aria-label="Mois des documents" sx={{ p: 2, mb: 2, position: 'sticky', top: 0, zIndex: 1, display: 'flex', flexWrap: 'wrap', alignItems: 'center', gap: 1.5, border: 1, borderColor: 'divider' }}>
      <PeriodSelector month={month} year={year} disabled={busy} onChange={setPeriod} />
      <Typography variant="caption" color="text.secondary">Mois de dépôt et de consultation</Typography>
    </Paper>
    {owners.error ? <Alert severity="error" sx={{ mb: 2 }}>{owners.error.message}<Button onClick={() => void owners.refetch()}>Réessayer</Button></Alert>
      : owners.isPending ? <CircularProgress size={22} /> : <DocumentUpload token={token} username={session.username} admin={admin} users={owners.data.users} month={month} year={year} busy={busy} run={run} />}
    <Paper sx={{ p: 2, my: 2 }}>
      <Box sx={{ display: 'flex', alignItems: 'center', flexWrap: 'wrap', gap: 2 }}>
        <Box sx={{ flex: 1 }}><Typography variant="h6">{archived ? 'Archives' : 'Tous les justificatifs'}</Typography><Typography variant="body2" color="text.secondary">{visible.length} document(s) correspondant aux filtres</Typography></Box>
        <Button disabled={busy || list.isFetching} onClick={() => void list.refetch()}>Actualiser</Button>
      </Box>
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
        <Button variant="contained" disabled={busy || !selection.length || selection.length > 100 || totalBytes > 100 * 1024 * 1024 || !!list.error} onClick={() => void run(async () => downloadFile(await documentsApi.export(token, selection.map(r => r.id))), 'ZIP prêt à télécharger.')}>Télécharger le ZIP ({selection.length})</Button>
        <Typography variant="caption" color="text.secondary">100 fichiers / 100 Mo maximum · {(totalBytes / 1024 / 1024).toFixed(1)} Mo sélectionnés</Typography>
      </Box>
    </Paper>
    {error && <Alert severity="error" sx={{ mb: 2 }} onClose={() => setError('')}>{error}</Alert>}
    {list.error ? <Alert severity="error">{list.error.message}<Button onClick={() => void list.refetch()}>Réessayer</Button></Alert> : list.isPending ? <CircularProgress size={28} /> :
      <Box sx={{ display: 'grid', gridTemplateColumns: { xs: '1fr', md: 'repeat(2, minmax(0, 1fr))', xl: 'repeat(3, minmax(0, 1fr))' }, gap: 2 }}>
        {!visible.length && <Typography color="text.secondary">Aucun document dans ce dossier pour ce mois.</Typography>}
        {visible.map(r => <DocumentCard key={r.id} receipt={r} admin={admin} canDelete={admin || (!manager && r.uploaded_by === session.username)} selected={selected.includes(r.id)} busy={busy}
          onSelect={() => setSelected(previous => previous.includes(r.id) ? previous.filter(id => id !== r.id) : [...previous, r.id])}
          onDownload={() => void run(async () => downloadFile(await documentsApi.file(token, r.id)), 'Document prêt à télécharger.')}
          onArchive={() => void run(async () => { await documentsApi.archive(token, r.id) }, 'Document déplacé dans les archives.')}
          onDelete={() => run(async () => { await documentsApi.remove(token, r.id) }, 'Document supprimé.')}
          onPeriod={(m, y) => run(async () => { await documentsApi.period(token, r.id, m, y) }, 'Mois de rattachement mis à jour.')} />)}
      </Box>}
    <Snackbar open={!!message} autoHideDuration={5000} onClose={() => setMessage('')}><Alert severity="success" variant="filled" onClose={() => setMessage('')}>{message}</Alert></Snackbar>
  </Box>
}
