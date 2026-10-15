import { alpha } from '@mui/material/styles'
import { useEffect, useRef, useState } from 'react'
import { Alert, Box, Button, Checkbox, CircularProgress, Dialog, DialogActions, DialogContent, DialogTitle, FormControlLabel, MenuItem, Paper, Snackbar, Tab, Tabs, TextField, Typography } from '@mui/material'
import DownloadRounded from '@mui/icons-material/DownloadRounded'
import SaveOutlined from '@mui/icons-material/SaveOutlined'
import { useQuery } from '@tanstack/react-query'
import type { Session } from '../types'
import { attendanceApi, downloadFile } from './api'
import { ficheLabel } from './types'
import type { Cell, Config, Fiche, Workspace } from './types'
import { PeriodSelector } from './PeriodSelector'
import { FicheEditor } from './FicheEditor'
import { AttendanceCalendar } from './AttendanceCalendar'
import { SignaturePreview } from './SignaturePreview'
import { ApiError } from '../api'
import { AccountOption } from './AccountOption'
import { accountOptionLabel } from './types'
import { FicheExportMenu } from './FicheExportMenu'
import { DeleteAction } from '../components/DeleteAction'
import { useDraftGuard } from '../draftGuard'
import { BatchExportResult } from '../components/BatchExportResult'
import type { ExportResult } from '../components/BatchExportResult'
import { periodLabel } from '../documents/api'

interface EditorState { dirty: boolean; busy: boolean }

export function AttendancePage({ token, session, onState }: { token: string; session: Session; onState: (state: EditorState) => void }) {
  useEffect(() => () => onState({ dirty: false, busy: false }), [onState])
  const [epoch, setEpoch] = useState(0)
  const workspace = useQuery({ queryKey: ['attendance-workspace', session.username], queryFn: ({ signal }) => attendanceApi.workspace(token, signal), retry: false })
  if (workspace.isPending) return <Box className="workspace"><CircularProgress size={28} /><Typography sx={{ mt: 2 }}>Chargement des fiches enregistrées…</Typography></Box>
  if (workspace.error && !workspace.data) return <Box className="workspace"><Alert severity="error">{workspace.error.message}</Alert><Button onClick={() => void workspace.refetch()}>Réessayer</Button></Box>
  return workspace.data ? <AttendanceEditor key={`${session.username}-${epoch}`} token={token} session={session} workspace={workspace.data} onState={onState} reload={async () => {
    const result = await workspace.refetch()
    if (result.error) throw result.error
    setEpoch(value => value + 1)
  }} /> : null
}

function AttendanceEditor({ token, session, workspace, onState, reload }: { token: string; session: Session; workspace: Workspace; onState: (state: EditorState) => void; reload: () => Promise<void> }) {
  const [config, setConfig] = useState<Config>(() => structuredClone(workspace.config))
  const [month, setMonth] = useState(() => workspace.config.mois || new Date().getMonth() + 1)
  const [year, setYear] = useState(() => workspace.config.annee || new Date().getFullYear())
  const [account, setAccount] = useState('')
  const [sheet, setSheet] = useState<'calendar' | 'identity' | 'planning'>('calendar')
  const [active, setActive] = useState<Record<string, string>>(() => {
    try { return JSON.parse(sessionStorage.getItem(`attendance-active-${session.username}`) || '{}') } catch { return {} }
  })
  const [selected, setSelected] = useState<Record<string, string[]>>({})
  const [dirty, setDirty] = useState(false)
  const [busy, setBusy] = useState(false)
  const [pendingCalendar, setPendingCalendar] = useState(false)
  const protectSelection = useDraftGuard(pendingCalendar, busy, 'Sélection du calendrier à appliquer')
  const operation = useRef(false)
  const revision = useRef(workspace.revision)
  const [conflict, setConflict] = useState(false)
  const [reloadConfirm, setReloadConfirm] = useState(false)
  useEffect(() => { onState({ dirty, busy }) }, [dirty, busy, onState])
  const [error, setError] = useState('')
  const [message, setMessage] = useState('')
  const [version, setVersion] = useState(0)
  const [format, setFormat] = useState<'original' | 'pdf'>('original')
  const [exportResult, setExportResult] = useState<ExportResult | null>(null)
  const [deleteConfirm, setDeleteConfirm] = useState(false)
  const [signing, setSigning] = useState<{ ids: string[]; month: number; year: number } | null>(null)
  useEffect(() => { try { sessionStorage.setItem(`attendance-active-${session.username}`, JSON.stringify(active)) } catch { /* Browsing still works without storage. */ } }, [active, session.username])
  const period = `${year}-${month}`
  const visible = config.employes_data.filter(e => !e.contract_id || (e.sheet_month === month && e.sheet_year === year))
  const current = visible.find(e => String(e.id) === active[period]) || visible[0]
  const selection = (selected[period] || []).filter(id => visible.some(e => String(e.id) === id))
  function setPeriod(m: number, y: number) { protectSelection(() => { setMonth(m); setYear(y); setDeleteConfirm(false) }) }
  function update(values: Partial<Fiche>) {
    if (!current) return
    setConfig(previous => ({ ...previous, employes_data: previous.employes_data.map(e => e.id === current.id ? { ...e, ...values } : e) }))
    setDirty(true)
  }
  async function persist() {
    if (config.employes_data.some(fiche => ['ddc', 'fdc', 'dds', 'fds'].some(key => fiche[key] === 'Invalid Date'))) {
      throw new Error('Veuillez corriger les dates invalides dans les informations personnelles et le contrat avant de sauvegarder.')
    }
    const data = { ...config, mois: month, annee: year }
    const next = await attendanceApi.save(token, data, revision.current)
    revision.current = next.revision
    setConfig(next.config); setDirty(false); setVersion(v => v + 1)
    return next
  }
  async function perform(action: () => Promise<void>) {
    if (operation.current) return
    operation.current = true
    setBusy(true); setError('')
    try { await action(); return true } catch (e) { if (e instanceof ApiError && e.status === 409 && e.message.toLowerCase().includes('configuration')) setConflict(true); setError(e instanceof Error ? e.message : 'Une erreur est survenue.'); return false }
    finally { operation.current = false; setBusy(false) }
  }
  async function exportSheets(ids: string[], chosen: 'original' | 'pdf') {
    if (operation.current) return
    const result: ExportResult = { phase: 'preparing', title: `Export ${chosen === 'pdf' ? 'PDF' : 'Excel / Word'} · ${periodLabel(month, year)}`,
      items: ids.map(id => { const fiche = config.employes_data.find(item => String(item.id) === id); return fiche ? ficheLabel(fiche) : `Fiche ${id}` }) }
    setExportResult(result)
    await perform(async () => {
      try {
      await persist()
      const file = await attendanceApi.export(token, ids, month, year, chosen)
      downloadFile(file)
      setExportResult({ ...result, phase: 'ready', filename: file.filename })
      setMessage(`${ids.length} fiche(s) préparée(s). Téléchargement déclenché.`)
      } catch (error) {
        setExportResult({ ...result, phase: 'error', error: error instanceof Error ? error.message : 'Export indisponible.' })
        throw error
      }
    })
  }
  async function apply(cells: Cell[], status: string, reason: string) {
    return perform(async () => {
      const saved = await persist()
      const next = await attendanceApi.apply(token, String(current.id), month, year, cells, status, reason, saved.revision)
      revision.current = next.revision
      setConfig(next.config); setDirty(false); setVersion(v => v + 1); setMessage('Présences enregistrées.')
    })
  }
  return <Box className="workspace attendance-workspace">
    <Box className="page-heading"><Box><Typography variant="overline" sx={{ letterSpacing: 2, color: 'primary.main', fontWeight: 700 }}>CRÉATION DE FICHES</Typography>
      <Typography variant="h4" component="h1">Générateur de fiche de présence</Typography>
      <Typography color="text.secondary" sx={{ mt: 1 }}>Retrouvez vos fiches, renseignez les présences et générez vos documents.</Typography></Box>
    </Box>
    <Paper className="attendance-controls"><PeriodSelector month={month} year={year} disabled={busy} onChange={setPeriod} />
      <Box className="attendance-person-controls">
        <TextField select label="Sélectionner l’employé existant" value={account} disabled={busy} onChange={e => setAccount(e.target.value)} sx={{ flex: 1, minWidth: 0 }} slotProps={{ select: {
          renderValue: value => { const selectedAccount = workspace.accounts.find(a => a.username === value); return selectedAccount ? <AccountOption account={selectedAccount} compact /> : '' },
          MenuProps: { slotProps: { paper: { sx: { maxHeight: 420, maxWidth: 'calc(100vw - 24px)' } } } },
        } }}>
          <MenuItem value="">Sélectionner une personne</MenuItem>{workspace.accounts.map(a => <MenuItem key={a.username} value={a.username} aria-label={accountOptionLabel(a)} sx={{ py: 1, minHeight: 60 }}><AccountOption account={a} /></MenuItem>)}
        </TextField>
        <Button variant="contained" disabled={!account || busy} onClick={() => protectSelection(() => void perform(async () => {
          if (dirty) await persist()
          const result = await attendanceApi.open(token, account, month, year, revision.current)
          revision.current = result.revision
          setConfig(result.config); setDirty(false); setActive(a => ({ ...a, [period]: result.active_id })); setMessage('Fiche ouverte.'); setVersion(v => v + 1)
        }))}>Ouvrir / créer sa fiche</Button>
      </Box>
    </Paper>
    {error && <Alert severity="error" onClose={() => setError('')} sx={{ my: 2 }}>{error}</Alert>}
    {conflict && <Alert severity="warning" sx={{ my: 2 }}>Le serveur contient une autre version. Votre brouillon reste ici.
      <Button disabled={busy} onClick={() => downloadFile({ blob: new Blob([JSON.stringify({ ...config, mois: month, annee: year }, null, 2)], { type: 'application/json' }), filename: `brouillon_${year}_${month}.json` })}>Télécharger mon brouillon</Button>
      <Button disabled={busy} onClick={() => setReloadConfirm(true)}>Recharger le serveur</Button>
    </Alert>}
    {busy && <Alert icon={<CircularProgress size={18} />} severity="info" sx={{ my: 2 }}>Traitement en cours…</Alert>}
    {exportResult && <BatchExportResult result={exportResult} />}
    {!current && <Paper className="requests-panel"><Box className="empty-state"><Typography variant="h6">Aucune fiche pour ce mois</Typography><Typography color="text.secondary">Sélectionnez un compte existant pour ouvrir sa fiche liée au contrat.</Typography></Box></Paper>}
    {current && <>
      <Tabs value={String(current.id)} onChange={(_, id: string) => { if (!busy && id !== String(current.id)) protectSelection(() => { setActive(a => ({ ...a, [period]: id })); setDeleteConfirm(false) }) }} variant="scrollable" scrollButtons="auto" sx={{ mb: 2 }}>
        {visible.map(e => <Tab key={String(e.id)} value={String(e.id)} label={ficheLabel(e)} />)}
      </Tabs>
      <Paper className="fiche-panel">
        <Box className="fiche-heading"><Box><Typography variant="h6" sx={{ display: 'flex', alignItems: 'baseline', flexWrap: 'wrap', gap: 1 }}>Fiche de <Box component="span" sx={theme => ({ bgcolor: alpha(theme.palette.primary.main, theme.palette.mode === 'dark' ? .14 : .09), color: theme.palette.primary.main, border: '1px solid', borderColor: alpha(theme.palette.primary.main, .3), borderRadius: 1, px: 1.25, py: .35, overflowWrap: 'anywhere', minWidth: 0 })}>{ficheLabel(current)}</Box></Typography><Typography variant="caption" color="text.secondary">{current.type || 'Salarié'} · {dirty ? 'Modifications non enregistrées' : 'Enregistré'}</Typography></Box>
          <Box sx={{ display: 'flex', gap: 1, flexWrap: 'wrap' }}>
            <Button variant="outlined" disabled={busy} onClick={() => void perform(async () => { await persist(); setSigning({ ids: [String(current.id)], month, year }) })}>Préparer la signature</Button>
            <FicheExportMenu key={`${period}-${current.id}`} intern={current.type === 'Stagiaire'} busy={busy} onExport={chosen => void exportSheets([String(current.id)], chosen)} />
            {workspace.can_edit_identity && <DeleteAction label="Supprimer cette fiche" disabled={busy} onClick={() => setDeleteConfirm(true)} />}
          </Box>
        </Box>
        {deleteConfirm && <Alert severity="warning" sx={{ mb: 2 }} action={<Button color="inherit" disabled={busy} onClick={() => void perform(async () => {
          const next = { ...config, mois: month, annee: year, employes_data: config.employes_data.filter(e => e.id !== current.id) }
          const saved = await attendanceApi.save(token, next, revision.current)
          revision.current = saved.revision
          setConfig(saved.config); setDirty(false); setDeleteConfirm(false); setMessage('Fiche supprimée.')
        })}>Confirmer</Button>}>Supprimer cette fiche ? <Button color="inherit" onClick={() => setDeleteConfirm(false)}>Annuler</Button></Alert>}
        <Tabs value={sheet} onChange={(_, value) => setSheet(value)} variant="scrollable" scrollButtons="auto" aria-label="Sections de la fiche" sx={{ mb: 3, borderBottom: 1, borderColor: 'divider' }}>
          <Tab value="calendar" id="fiche-tab-calendar" aria-controls="fiche-panel-calendar" label="Calendrier mensuel" />
          <Tab value="identity" id="fiche-tab-identity" aria-controls="fiche-panel-identity" label="Informations personnelles et contrat" />
          <Tab value="planning" id="fiche-tab-planning" aria-controls="fiche-panel-planning" label={current.type === 'Stagiaire' ? 'Rémunération et présence' : 'Planning hebdomadaire'} />
        </Tabs>
        <Box role="tabpanel" id="fiche-panel-calendar" aria-labelledby="fiche-tab-calendar" hidden={sheet !== 'calendar'}>
          <AttendanceCalendar key={`${current.id}-${period}`} token={token} id={String(current.id)} month={month} year={year} version={version} busy={busy} active={sheet === 'calendar'} onPeriod={setPeriod} onApply={apply} onSelectionPending={setPendingCalendar} />
        </Box>
        {(['identity', 'planning'] as const).map(section => <Box key={section} role="tabpanel" id={`fiche-panel-${section}`} aria-labelledby={`fiche-tab-${section}`} hidden={sheet !== section}>
          <FicheEditor key={String(current.id)} fiche={current} section={section} admin={workspace.can_edit_identity} disabled={busy} onChange={update} />
        </Box>)}
      </Paper>
      <Paper className="batch-panel">
        <Typography variant="h6">Export de plusieurs fiches</Typography><Typography variant="body2" color="text.secondary" sx={{ mt: 1 }}>Choisissez les personnes à inclure dans le ZIP de ce mois.</Typography>
        <Box sx={{ display: 'flex', flexWrap: 'wrap', gap: 1, mt: 2 }}>
          <Button disabled={busy} onClick={() => setSelected(s => ({ ...s, [period]: visible.map(e => String(e.id)) }))}>Tout sélectionner</Button>
          <Button disabled={busy} onClick={() => setSelected(s => ({ ...s, [period]: [] }))}>Tout désélectionner</Button>
        </Box>
        <Box className="batch-selection">{visible.map(e => <FormControlLabel key={String(e.id)} label={ficheLabel(e)} control={<Checkbox disabled={busy} checked={selection.includes(String(e.id))}
          onChange={event => setSelected(s => ({ ...s, [period]: event.target.checked ? [...selection, String(e.id)] : selection.filter(id => id !== String(e.id)) }))} />} />)}</Box>
        <Box className="calendar-toolbar"><TextField select label="Format" value={format} disabled={busy} onChange={e => setFormat(e.target.value as 'original' | 'pdf')} sx={{ minWidth: 170 }}>
          <MenuItem value="original">Excel / Word</MenuItem><MenuItem value="pdf">PDF</MenuItem>
        </TextField><Button variant="contained" disabled={!selection.length || busy} startIcon={<DownloadRounded />} onClick={() => void exportSheets(selection, format)}>Générer la sélection ({selection.length})</Button>
        <Button variant="outlined" disabled={!selection.length || busy} onClick={() => void perform(async () => { await persist(); setSigning({ ids: [...selection], month, year }) })}>Préparer les signatures ({selection.length})</Button></Box>
      </Paper>
    </>}
    <Paper className="attendance-floating-save" sx={{ zIndex: theme => theme.zIndex.appBar }}>
      <Typography variant="caption" color={dirty ? 'primary.main' : 'text.secondary'} sx={{ display: 'block', textAlign: 'center', mb: .5 }}>
        {busy ? 'Traitement en cours…' : dirty ? 'Modifications à enregistrer' : pendingCalendar ? 'Sélection à appliquer' : 'Aucune modification en attente'}
      </Typography>
      <Button variant={dirty ? 'contained' : 'outlined'} startIcon={<SaveOutlined />} disabled={busy || (!dirty && pendingCalendar)} onClick={() => void perform(async () => { await persist(); setMessage('Configuration sauvegardée.') })}>{dirty ? 'Enregistrer les modifications' : 'Sauvegarder'}</Button>
    </Paper>
    {signing && <SignaturePreview token={token} {...signing} onClose={() => setSigning(null)} />}
    <Dialog open={reloadConfirm} onClose={() => { if (!busy) setReloadConfirm(false) }}><DialogTitle>Recharger la version du serveur ?</DialogTitle><DialogContent>Le brouillon affiché sera remplacé. Téléchargez-le d’abord si vous souhaitez le conserver.</DialogContent><DialogActions><Button disabled={busy} onClick={() => setReloadConfirm(false)}>Annuler</Button><Button disabled={busy} onClick={() => void perform(reload)}>Recharger</Button></DialogActions></Dialog>
    <Snackbar open={Boolean(message)} autoHideDuration={4500} onClose={() => setMessage('')} sx={{ bottom: 'calc(88px + env(safe-area-inset-bottom))' }}><Alert severity="success" variant="filled" onClose={() => setMessage('')}>{message}</Alert></Snackbar>
  </Box>
}
