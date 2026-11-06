import { useCallback, useState } from 'react'
import { Alert, Autocomplete, Box, Button, Chip, CircularProgress, Dialog, DialogActions, DialogContent, DialogTitle, IconButton, MenuItem, Paper, Tab, Table, TableBody, TableCell, TableContainer, TableHead, TablePagination, TableRow, Tabs, TextField, Tooltip, Typography } from '@mui/material'
import AddRounded from '@mui/icons-material/AddRounded'
import HistoryRounded from '@mui/icons-material/HistoryRounded'
import UploadFileOutlined from '@mui/icons-material/UploadFileOutlined'
import { useQuery, useQueryClient } from '@tanstack/react-query'
import dayjs from 'dayjs'
import type { Session } from '../types'
import { isAdmin, isSuperAdmin } from '../rolePolicy'
import { WorkspaceToolbar } from '../components/WorkspaceToolbar'
import { FrenchDateField } from '../components/FrenchDateField'
import { DocumentPreview } from '../documents/DocumentPreview'
import { useDraftGuard } from '../draftGuard'
import { medicalApi, readPdf } from './api'
import type { MedicalEmployee, MedicalVisit } from './api'
import { MedicalHistory } from './MedicalHistory'
import { MedicalExport } from './MedicalExport'
import { linkedMedical } from './medicalLink'

const dateLabel = (date: string | null) => date ? dayjs(date).format('DD/MM/YYYY') : '—'

export function MedicalVisitsPage({ token, session }: { token: string; session: Session }) {
  const cache = useQueryClient(), admin = isAdmin(session)
  const [q, setQ] = useState(''), [page, setPage] = useState(0), [archived, setArchived] = useState(false)
  const [year, setYear] = useState(dayjs().year()), [month, setMonth] = useState<number | null>(null)
  const period = { year, month }
  const [person, setPerson] = useState<MedicalEmployee | null>(null)
  const [form, setForm] = useState<{ person?: MedicalEmployee; visit?: MedicalVisit } | null>(null)
  const [preview, setPreview] = useState<MedicalVisit | null>(null)
  const people = useQuery({ queryKey: ['medical', session.username, 'employees', q, page, archived, year, month], queryFn: ({ signal }) => medicalApi.employees(token, q, page + 1, archived, signal, period), retry: false })
  const own = useQuery({ queryKey: ['medical', session.username, 'self'], queryFn: ({ signal }) => medicalApi.employees(token, '', 1, false, signal, undefined, true), enabled: !admin, retry: false })
  const years = [...new Set([year, dayjs().year(), ...(people.data?.years || []), ...Array.from({ length: 10 }, (_, i) => dayjs().year() - i)])].sort((a,b) => b-a)
  const [linkedId, setLinkedId] = useState(() => linkedMedical(window.location.search))
  const linked = useQuery({ queryKey: ['medical', session.username, 'linked', linkedId], queryFn: ({ signal }) => medicalApi.record(token, linkedId!, signal), enabled: linkedId !== null, retry: false })
  const visiblePreview = preview || (linkedId ? linked.data : null)
  const loadFile = useCallback((signal: AbortSignal) => medicalApi.file(token, visiblePreview!.id, signal, visiblePreview?.revision), [token, visiblePreview])
  const saved = async () => { setForm(null); await cache.invalidateQueries({ queryKey: ['medical'] }) }
  return <Box className="workspace" sx={{ maxWidth: 'none', '& .MuiTableCell-root': { py: 1, fontSize: 13 } }}>
    <WorkspaceToolbar>
      <TextField select size="small" label="Année" value={year} onChange={e => { setYear(Number(e.target.value)); setPage(0); setPerson(null) }} sx={{ width: 100 }}>{years.map(value => <MenuItem key={value} value={value}>{value}</MenuItem>)}</TextField>
      <TextField select size="small" label="Mois" value={month || 'all'} onChange={e => { setMonth(Number(e.target.value) || null); setPage(0); setPerson(null) }} sx={{ width: 150 }}><MenuItem value="all">Tous les mois</MenuItem>{Array.from({ length: 12 }, (_, i) => <MenuItem key={i+1} value={i+1}>{new Intl.DateTimeFormat('fr-FR', { month: 'long' }).format(new Date(2026, i, 1))}</MenuItem>)}</TextField>
      <TextField size="small" placeholder="Rechercher une personne" slotProps={{ htmlInput: { 'aria-label': 'Rechercher une personne' } }} value={q} onChange={e => { setQ(e.target.value); setPage(0) }} sx={{ width: { xs: 150, sm: 220 } }} />
      <Tabs value={archived ? 1 : 0} onChange={(_, value) => { setArchived(value === 1); setPage(0) }} aria-label="Dossiers" sx={{ '& .MuiTab-root': { minWidth: 76, px: 1 } }}><Tab label="Normal" /><Tab label="Archives" /></Tabs>
      <MedicalExport token={token} period={period} q={q} archived={archived} disabled={!people.data?.total} />
      <Button className="workspace-primary-action" variant="contained" startIcon={<AddRounded />} disabled={archived || (!admin && !own.data?.total)} onClick={() => setForm({ person: !admin ? own.data?.employees[0] : undefined })}>Ajouter une visite</Button>
    </WorkspaceToolbar>
    {people.error && <Alert severity="error">{people.error.message}</Alert>}
    {linked.error && <Alert severity="error" onClose={() => setLinkedId(null)}>{linked.error.message}</Alert>}
    {linkedId && linked.data && !linked.data.original_filename && <Alert severity="info" onClose={() => setLinkedId(null)}>Attestation non déposée.</Alert>}
    <Paper variant="outlined" sx={{ overflow: 'hidden' }}>
      <TableContainer><Table aria-label="Visites médicales" sx={{ tableLayout: 'fixed' }}><TableHead><TableRow>
        <TableCell sx={{ width: '28%' }}>Nom et prénoms</TableCell><TableCell sx={{ display: { xs: 'none', sm: 'table-cell' } }}>Département</TableCell><TableCell>Dernière visite</TableCell><TableCell>Attestation</TableCell><TableCell sx={{ width: 102 }} />
      </TableRow></TableHead><TableBody>
        {people.data?.employees.map(row => <TableRow key={row.employee_id} hover>
          <TableCell><Button sx={{ justifyContent: 'flex-start', p: 0, textAlign: 'left', overflowWrap: 'anywhere' }} onClick={() => setPerson(row)}>{row.full_name}</Button></TableCell>
          <TableCell sx={{ display: { xs: 'none', sm: 'table-cell' }, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }} title={row.departments.join(', ')}>{row.departments.join(', ') || '—'}</TableCell>
          <TableCell>{dateLabel(row.visit_date)}</TableCell><TableCell>{row.original_filename ? <Chip label="Déposée" size="small" variant="outlined" /> : <Typography variant="body2" color="text.secondary">{row.latest_id ? 'À déposer' : '—'}</Typography>}</TableCell>
          <TableCell><Box sx={{ display: 'flex', justifyContent: 'flex-end', alignItems: 'center' }}><MedicalExport token={token} period={period} archived={archived} employeeId={row.employee_id} allHistory name={row.full_name} /><Tooltip title="Voir les visites"><IconButton aria-label={`Visites de ${row.full_name}`} onClick={() => setPerson(row)}><HistoryRounded fontSize="small" /></IconButton></Tooltip></Box></TableCell>
        </TableRow>)}
      </TableBody></Table></TableContainer>
      {people.isLoading && <Box sx={{ p: 3, textAlign: 'center' }}><CircularProgress size={24} /></Box>}
      {!people.isLoading && !people.error && !people.data?.total && <Typography color="text.secondary" sx={{ p: 2 }}>{!admin && own.isSuccess && !own.data.total ? 'Aucun dossier salarié associé. Contactez un administrateur.' : 'Aucune visite pour cette période.'}</Typography>}
      <TablePagination component="div" count={people.data?.total || 0} page={page} onPageChange={(_, value) => setPage(value)} rowsPerPage={20} rowsPerPageOptions={[20]} labelDisplayedRows={({ from, to, count }) => `${from}–${to} sur ${count}`} />
    </Paper>
    {person && <MedicalHistory key={`${person.employee_id}-${year}-${month}-${archived}`} token={token} username={session.username} admin={admin} superAdmin={isSuperAdmin(session)} archived={archived} period={period} person={person} onClose={() => setPerson(null)} onEdit={visit => setForm({ person, visit })} onPreview={setPreview} />}
    {form && <VisitForm token={token} username={session.username} admin={admin} initial={form.person} visit={form.visit} onClose={() => setForm(null)} onSaved={saved} />}
    {visiblePreview?.original_filename && <DocumentPreview token={token} receipt={{ id: visiblePreview.id, original_filename: visiblePreview.original_filename!, document_number: null }} loadFile={loadFile} onClose={() => { setPreview(null); setLinkedId(null); const url = new URL(window.location.href); url.searchParams.delete('medical'); window.history.replaceState(null, '', url) }} />}
  </Box>
}

function VisitForm({ token, username, admin, initial, visit, onClose, onSaved }: { token: string; username: string; admin: boolean; initial?: MedicalEmployee; visit?: MedicalVisit; onClose: () => void; onSaved: () => Promise<void> }) {
  const [person, setPerson] = useState(initial || null), [search, setSearch] = useState('')
  const [date, setDate] = useState(visit?.visit_date || ''), [file, setFile] = useState<File | null>(null)
  const [busy, setBusy] = useState(false), [error, setError] = useState('')
  const close = useDraftGuard(!!file || (visit ? date !== visit.visit_date : (!!date || person?.employee_id !== initial?.employee_id)), busy, 'Visite médicale')
  const choices = useQuery({ queryKey: ['medical', username, 'choices', search], queryFn: ({ signal }) => medicalApi.employees(token, search, 1, false, signal), enabled: admin && !initial, retry: false })
  async function submit() {
    if (!person || !dayjs(date, 'YYYY-MM-DD', true).isValid() || dayjs(date).isAfter(dayjs(), 'day')) { setError('Choisissez une personne et une date de visite valide.'); return }
    setBusy(true); setError('')
    try {
      const pdf = file ? await readPdf(file) : null
      if (visit) await medicalApi.correct(token, visit, date, pdf)
      else await medicalApi.create(token, person.employee_id, date, pdf)
      await onSaved()
    } catch (reason) { setError(reason instanceof Error ? reason.message : 'Enregistrement impossible.') }
    finally { setBusy(false) }
  }
  return <Dialog open fullWidth maxWidth="sm" onClose={() => { if (!busy) close(onClose) }}>
    <DialogTitle>{visit ? 'Corriger la visite' : 'Enregistrer une visite'}</DialogTitle>
    <DialogContent dividers>
      {initial ? <Typography sx={{ mb: 2 }}>{initial.full_name}</Typography> : <Autocomplete options={choices.data?.employees || []} value={person} getOptionLabel={item => item.full_name} isOptionEqualToValue={(a,b) => a.employee_id === b.employee_id} filterOptions={items => items} onChange={(_, value) => setPerson(value)} onInputChange={(_, value) => setSearch(value)} disabled={busy} renderInput={params => <TextField {...params} label="Salarié" required />} sx={{ mb: 2 }} />}
      {choices.error && <Alert severity="error">{choices.error.message}</Alert>}
      <FrenchDateField label="Date de visite" value={date} onChange={setDate} required fullWidth disabled={busy} />
      <Box sx={{ display: 'flex', gap: 2, alignItems: 'center', flexWrap: 'wrap', mt: 2 }}><Button component="label" variant="outlined" disabled={busy} startIcon={<UploadFileOutlined />}>Choisir le PDF<input type="file" hidden accept="application/pdf,.pdf" onChange={e => setFile(e.target.files?.[0] || null)} /></Button><Typography variant="caption" color="text.secondary">PDF · 10 Mo max.{' · facultatif'}</Typography></Box>
      {visit?.original_filename && !file && <Typography variant="caption" color="text.secondary" sx={{ display: 'block', mt: 1 }}>PDF actuel conservé sans nouveau fichier.</Typography>}
      {file && <Typography variant="body2" sx={{ mt: 1, overflowWrap: 'anywhere' }}>{file.name}</Typography>}
      {error && <Alert severity="error" sx={{ mt: 2 }}>{error}</Alert>}
    </DialogContent>
    <DialogActions><Button disabled={busy} onClick={() => close(onClose)}>Annuler</Button><Button variant="contained" disabled={busy || !person || !date} onClick={() => void submit()}>{busy ? 'Enregistrement…' : 'Enregistrer'}</Button></DialogActions>
  </Dialog>
}
