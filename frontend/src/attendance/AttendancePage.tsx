import { alpha } from '@mui/material/styles'
import { useEffect, useRef, useState } from 'react'
import { Alert, Box, Button, Checkbox, CircularProgress, Dialog, DialogActions, DialogContent, DialogTitle, FormControlLabel, MenuItem, Paper, Snackbar, Tab, Tabs, TextField, Tooltip, Typography } from '@mui/material'
import DownloadRounded from '@mui/icons-material/DownloadRounded'
import SaveOutlined from '@mui/icons-material/SaveOutlined'
import { useQuery, useQueryClient } from '@tanstack/react-query'
import type { Session } from '../types'
import { attendanceApi, downloadFile } from './api'
import { accountSheetPeriods, attendancePersonName, ficheLabel, withAttendanceLabels } from './types'
import type { Cell, Config, Fiche, SavedConfig, Workspace } from './types'
import { FicheEditor } from './FicheEditor'
import { AttendanceCalendar } from './AttendanceCalendar'
import { SignaturePreview } from './SignaturePreview'
import { ApiError } from '../api'
import { AttendanceToolbar } from './AttendanceToolbar'
import { FicheExportMenu } from './FicheExportMenu'
import { DeleteAction } from '../components/DeleteAction'
import { PermanentDeleteDialog } from '../components/PermanentDeleteDialog'
import ArchiveOutlined from '@mui/icons-material/ArchiveOutlined'
import UnarchiveOutlined from '@mui/icons-material/UnarchiveOutlined'
import { isAdmin, isSuperAdmin } from '../rolePolicy'
import { useDraftGuard } from '../draftGuard'
import { BatchExportResult } from '../components/BatchExportResult'
import type { ExportResult } from '../components/BatchExportResult'
import { periodLabel } from '../documents/api'
import { personDisplayName } from '../personDisplayName'

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
  const client = useQueryClient()
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
  const [deleteConfirm, setDeleteConfirm] = useState<Fiche | null>(null)
  const [archiveConfirm, setArchiveConfirm] = useState<Fiche | null>(null)
  const [archived, setArchived] = useState(false)
  const [signing, setSigning] = useState<{ ids: string[]; month: number; year: number } | null>(null)
  useEffect(() => { try { sessionStorage.setItem(`attendance-active-${session.username}`, JSON.stringify(active)) } catch { /* Browsing still works without storage. */ } }, [active, session.username])
  const period = `${year}-${month}`
  const visible = config.employes_data.filter(e => !!e.is_archived === archived && (!e.contract_id || (e.sheet_month === month && e.sheet_year === year)))
  const current = visible.find(e => String(e.id) === active[period]) || visible[0]
  const selection = (selected[period] || []).filter(id => visible.some(e => String(e.id) === id))
  const chosenAccount = workspace.accounts.find(person => person.username === account)
  const accountInMonth = visible.some(fiche => fiche.account_username === account)
  const savedPeriods = accountSheetPeriods(config.employes_data, account, archived)
  function personNameFor(fiche: Fiche) {
    return attendancePersonName(fiche, workspace)
  }
  function ficheDisplayLabel(fiche: Fiche) { return ficheLabel(fiche, personNameFor(fiche)) }
  function syncLabels(saved: SavedConfig) {
    if (saved.person_labels) client.setQueryData<Workspace>(['attendance-workspace', session.username], previous => previous ? withAttendanceLabels(previous, saved) : previous)
  }
  function setPeriod(m: number, y: number) { protectSelection(() => { setMonth(m); setYear(y); setDeleteConfirm(null) }) }
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
    revision.current = next.revision; syncLabels(next)
    setConfig(next.config); setDirty(false); setVersion(v => v + 1)
    return next
  }
  async function perform(action: () => Promise<void>, rethrow = false) {
    if (operation.current) { if (rethrow) throw new Error('Une autre opération est en cours.'); return }
    operation.current = true
    setBusy(true); setError('')
    try { await action(); return true } catch (e) { if (e instanceof ApiError && e.status === 409 && e.message.toLowerCase().includes('configuration')) setConflict(true); setError(e instanceof Error ? e.message : 'Une erreur est survenue.'); if (rethrow) throw e; return false }
    finally { operation.current = false; setBusy(false) }
  }
  async function exportSheets(ids: string[], chosen: 'original' | 'pdf') {
    if (operation.current) return
    const result: ExportResult = { phase: 'preparing', title: `Export ${chosen === 'pdf' ? 'PDF' : 'Excel / Word'} · ${periodLabel(month, year)}`,
      items: ids.map(id => { const fiche = config.employes_data.find(item => String(item.id) === id); return fiche ? ficheDisplayLabel(fiche) : `Fiche ${id}` }) }
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
      revision.current = next.revision; syncLabels(next)
      setConfig(next.config); setDirty(false); setVersion(v => v + 1); setMessage('Présences enregistrées.')
    })
  }
  return <Box className="workspace attendance-workspace" sx={{ '&.workspace': { maxWidth: 'none', px: { xs: 1.5, sm: 2 }, pt: 1.5 }, '& .fiche-panel, & .batch-panel': { p: { xs: 1.5, sm: 2 }, mb: 2 } }}>
    <AttendanceToolbar month={month} year={year} busy={busy} account={account} accounts={workspace.accounts} onPeriod={setPeriod} onAccount={value => protectSelection(() => {
      setAccount(value); setError('')
      const existing = visible.find(fiche => fiche.account_username === value)
      if (existing) setActive(previous => ({ ...previous, [period]: String(existing.id) }))
    })} archived={archived} onArchiveFilter={value => protectSelection(() => { setArchived(value); setSelected({}) })} onOpen={() => protectSelection(() => void perform(async () => {
          if (dirty) await persist()
          const result = await attendanceApi.open(token, account, month, year, revision.current)
          revision.current = result.revision; syncLabels(result)
          setConfig(result.config); setDirty(false); setArchived(false); setActive(a => ({ ...a, [period]: result.active_id })); setMessage('Fiche ouverte.'); setVersion(v => v + 1)
        }))} />
    {chosenAccount && !accountInMonth && <Alert severity="info" sx={{ mb: 2 }}>
      <Typography variant="body2">{personDisplayName(chosenAccount)} : aucune fiche {archived ? 'archivée ' : ''}enregistrée pour {periodLabel(month, year)}.{visible.length > 0 ? ' Les onglets ci-dessous correspondent aux autres personnes de ce mois.' : ''}</Typography>
      {savedPeriods.length > 0 && <Box sx={{ display: 'flex', alignItems: 'center', flexWrap: 'wrap', gap: .5, mt: .5 }}>
        <Typography variant="body2" sx={{ mr: .5 }}>Fiches existantes :</Typography>
        {savedPeriods.slice(0, 3).map(saved => <Button key={`${saved.year}-${saved.month}`} size="small" disabled={busy} onClick={() => protectSelection(() => {
          setMonth(saved.month); setYear(saved.year); setError(''); setDeleteConfirm(null)
          setActive(previous => ({ ...previous, [`${saved.year}-${saved.month}`]: saved.id }))
        })}>Voir {periodLabel(saved.month, saved.year)}</Button>)}
        {savedPeriods.length > 3 && <TextField select size="small" value="" label="Autres mois" sx={{ minWidth: 170 }} disabled={busy} onChange={event => {
          const saved = savedPeriods.find(item => `${item.year}-${item.month}` === event.target.value)
          if (saved) protectSelection(() => { setMonth(saved.month); setYear(saved.year); setError(''); setDeleteConfirm(null); setActive(previous => ({ ...previous, [`${saved.year}-${saved.month}`]: saved.id })) })
        }}><MenuItem value="">Choisir un mois</MenuItem>{savedPeriods.slice(3).map(saved => <MenuItem key={`${saved.year}-${saved.month}`} value={`${saved.year}-${saved.month}`}>{periodLabel(saved.month, saved.year)}</MenuItem>)}</TextField>}
      </Box>}
    </Alert>}
    {error && <Alert severity="error" onClose={() => setError('')} sx={{ my: 2 }}>{error}</Alert>}
    {conflict && <Alert severity="warning" sx={{ my: 2 }}>Le serveur contient une autre version. Votre brouillon reste ici.
      <Button disabled={busy} onClick={() => downloadFile({ blob: new Blob([JSON.stringify({ ...config, mois: month, annee: year }, null, 2)], { type: 'application/json' }), filename: `brouillon_${year}_${month}.json` })}>Télécharger mon brouillon</Button>
      <Button disabled={busy} onClick={() => setReloadConfirm(true)}>Recharger le serveur</Button>
    </Alert>}
    {busy && <Alert icon={<CircularProgress size={18} />} severity="info" sx={{ my: 2 }}>Traitement en cours…</Alert>}
    {exportResult && <BatchExportResult result={exportResult} />}
    {!current && <Paper className="requests-panel"><Box className="empty-state"><Typography variant="h6">Aucune fiche pour ce mois</Typography><Typography color="text.secondary">Sélectionnez un compte existant pour ouvrir sa fiche liée au contrat.</Typography></Box></Paper>}
    {current && <>
      <Tabs value={String(current.id)} onChange={(_, id: string) => { if (!busy && id !== String(current.id)) protectSelection(() => { setActive(a => ({ ...a, [period]: id })); setDeleteConfirm(null) }) }} variant="scrollable" scrollButtons="auto" sx={{ mb: 2 }}>
        {visible.map(e => <Tab key={String(e.id)} value={String(e.id)} label={ficheDisplayLabel(e)} />)}
      </Tabs>
      <Paper className="fiche-panel">
        <Box className="fiche-heading"><Box><Typography variant="h6" sx={{ display: 'flex', alignItems: 'baseline', flexWrap: 'wrap', gap: 1 }}>Fiche de <Box component="span" sx={theme => ({ bgcolor: alpha(theme.palette.primary.main, theme.palette.mode === 'dark' ? .14 : .09), color: theme.palette.primary.main, border: '1px solid', borderColor: alpha(theme.palette.primary.main, .3), borderRadius: 1, px: 1.25, py: .35, overflowWrap: 'anywhere', minWidth: 0 })}>{ficheDisplayLabel(current)}</Box></Typography><Typography variant="caption" color="text.secondary">{current.type || 'Salarié'} · {dirty ? 'Modifications non enregistrées' : 'Enregistré'}</Typography></Box>
          <Box sx={{ display: 'flex', gap: 1, flexWrap: 'wrap' }}>
            {!current.is_archived && <Button variant="outlined" disabled={busy} onClick={() => void perform(async () => { await persist(); setSigning({ ids: [String(current.id)], month, year }) })}>Préparer la signature</Button>}
            <FicheExportMenu key={`${period}-${current.id}`} intern={current.type === 'Stagiaire'} busy={busy} onExport={chosen => void exportSheets([String(current.id)], chosen)} />
            {isAdmin(session) && workspace.can_archive_sheets && <Button startIcon={current.is_archived ? <UnarchiveOutlined /> : <ArchiveOutlined />} disabled={busy} onClick={() => protectSelection(() => setArchiveConfirm({ ...current }))}>{current.is_archived ? 'Restaurer' : 'Archiver'}</Button>}
            {isSuperAdmin(session) && workspace.can_delete_sheets && current.is_archived && <DeleteAction label="Supprimer définitivement cette fiche" disabled={busy} onClick={() => protectSelection(() => setDeleteConfirm({ ...current }))} />}
          </Box>
        </Box>
        {current.is_archived && <Alert severity="info" sx={{ mb: 2 }}>Fiche archivée : consultation et export disponibles. Restaurez-la pour la modifier ou préparer une nouvelle signature.</Alert>}
        <Tabs value={sheet} onChange={(_, value) => setSheet(value)} variant="scrollable" scrollButtons="auto" aria-label="Sections de la fiche" sx={{ mb: 3, borderBottom: 1, borderColor: 'divider' }}>
          <Tab value="calendar" id="fiche-tab-calendar" aria-controls="fiche-panel-calendar" label="Calendrier mensuel" />
          <Tab value="identity" id="fiche-tab-identity" aria-controls="fiche-panel-identity" label="Informations personnelles et contrat" />
          <Tab value="planning" id="fiche-tab-planning" aria-controls="fiche-panel-planning" label={current.type === 'Stagiaire' ? 'Rémunération et présence' : 'Planning hebdomadaire'} />
        </Tabs>
        <Box role="tabpanel" id="fiche-panel-calendar" aria-labelledby="fiche-tab-calendar" hidden={sheet !== 'calendar'}>
          <AttendanceCalendar key={`${current.id}-${period}`} token={token} id={String(current.id)} month={month} year={year} version={version} busy={busy || !!current.is_archived} active={sheet === 'calendar'} showPeriodSelector={false} onPeriod={setPeriod} onApply={apply} onSelectionPending={setPendingCalendar} />
        </Box>
        {(['identity', 'planning'] as const).map(section => <Box key={section} role="tabpanel" id={`fiche-panel-${section}`} aria-labelledby={`fiche-tab-${section}`} hidden={sheet !== section}>
          <FicheEditor key={String(current.id)} fiche={current} section={section} admin={workspace.can_edit_identity} disabled={busy || !!current.is_archived} onChange={update} />
        </Box>)}
      </Paper>
      {!archived && <Paper className="batch-panel">
        <Tooltip title="Choisissez les personnes à inclure dans le ZIP de ce mois. Le téléchargement et la préparation des signatures sont deux opérations distinctes." describeChild enterDelay={650} arrow><Typography component="h2" variant="h6" tabIndex={0} sx={{ width: 'fit-content', cursor: 'help' }}>Export de plusieurs fiches</Typography></Tooltip>
        <Box sx={{ display: 'flex', flexWrap: 'wrap', gap: 1, mt: 2 }}>
          <Button disabled={busy} onClick={() => setSelected(s => ({ ...s, [period]: visible.map(e => String(e.id)) }))}>Tout sélectionner</Button>
          <Button disabled={busy} onClick={() => setSelected(s => ({ ...s, [period]: [] }))}>Tout désélectionner</Button>
        </Box>
        <Box className="batch-selection">{visible.map(e => <FormControlLabel key={String(e.id)} label={ficheDisplayLabel(e)} control={<Checkbox disabled={busy} checked={selection.includes(String(e.id))}
          onChange={event => setSelected(s => ({ ...s, [period]: event.target.checked ? [...selection, String(e.id)] : selection.filter(id => id !== String(e.id)) }))} />} />)}</Box>
        <Box className="calendar-toolbar"><TextField select label="Format" value={format} disabled={busy} onChange={e => setFormat(e.target.value as 'original' | 'pdf')} sx={{ minWidth: 170 }}>
          <MenuItem value="original">Excel / Word</MenuItem><MenuItem value="pdf">PDF</MenuItem>
        </TextField><Button variant="contained" disabled={!selection.length || busy} startIcon={<DownloadRounded />} onClick={() => void exportSheets(selection, format)}>Générer la sélection ({selection.length})</Button>
        <Button variant="outlined" disabled={!selection.length || busy} onClick={() => void perform(async () => { await persist(); setSigning({ ids: [...selection], month, year }) })}>Préparer les signatures ({selection.length})</Button></Box>
      </Paper>}
    </>}
    <Paper className="attendance-floating-save" sx={{ zIndex: theme => theme.zIndex.appBar }}>
      <Typography variant="caption" color={dirty ? 'primary.main' : 'text.secondary'} sx={{ display: 'block', textAlign: 'center', mb: .5 }}>
        {busy ? 'Traitement en cours…' : dirty ? 'Modifications à enregistrer' : pendingCalendar ? 'Sélection à appliquer' : 'Aucune modification en attente'}
      </Typography>
      <Button variant={dirty ? 'contained' : 'outlined'} startIcon={<SaveOutlined />} disabled={busy || (!dirty && pendingCalendar)} onClick={() => void perform(async () => { await persist(); setMessage('Configuration sauvegardée.') })}>{dirty ? 'Enregistrer les modifications' : 'Sauvegarder'}</Button>
    </Paper>
    {signing && <SignaturePreview token={token} {...signing} personLabels={Object.fromEntries(config.employes_data.map(fiche => [String(fiche.id), personNameFor(fiche)]))} onClose={() => setSigning(null)} />}
    {archiveConfirm && <Dialog open onClose={() => { if (!busy) setArchiveConfirm(null) }}><DialogTitle>{archiveConfirm.is_archived ? 'Restaurer cette fiche ?' : 'Archiver cette fiche ?'}</DialogTitle><DialogContent><Alert severity="info">{ficheDisplayLabel(archiveConfirm)} : la fiche et son historique seront conservés. Les modifications déjà saisies seront enregistrées en même temps.</Alert>{error && <Alert severity="error" sx={{ mt: 2 }}>{error}</Alert>}</DialogContent><DialogActions><Button disabled={busy} onClick={() => setArchiveConfirm(null)}>Annuler</Button><Button disabled={busy} onClick={() => void perform(async () => {
      const next = { ...config, mois: month, annee: year, employes_data: config.employes_data.map(fiche => fiche.id === archiveConfirm.id ? { ...fiche, is_archived: !archiveConfirm.is_archived } : fiche) }
      const saved = await attendanceApi.save(token, next, revision.current)
      revision.current = saved.revision; syncLabels(saved); setConfig(saved.config); setDirty(false); setVersion(value => value + 1); setArchiveConfirm(null); setMessage(archiveConfirm.is_archived ? 'Fiche restaurée.' : 'Fiche archivée.')
    })}>Confirmer</Button></DialogActions></Dialog>}
    {deleteConfirm && <PermanentDeleteDialog title="Supprimer définitivement cette fiche ?" expected={String(deleteConfirm.id)} description={`${ficheDisplayLabel(deleteConfirm)} sera supprimée. Les modifications déjà saisies sur les autres fiches seront enregistrées en même temps.`} onClose={() => setDeleteConfirm(null)} onConfirm={() => perform(async () => {
      const next = { ...config, mois: month, annee: year, employes_data: config.employes_data.filter(fiche => fiche.id !== deleteConfirm.id) }
      const saved = await attendanceApi.save(token, next, revision.current, [String(deleteConfirm.id)])
      revision.current = saved.revision; syncLabels(saved); setConfig(saved.config); setDirty(false); setVersion(value => value + 1)
    }, true)} onDeleted={() => { setDeleteConfirm(null); setMessage('Fiche supprimée.') }} />}
    <Dialog open={reloadConfirm} onClose={() => { if (!busy) setReloadConfirm(false) }}><DialogTitle>Recharger la version du serveur ?</DialogTitle><DialogContent>Le brouillon affiché sera remplacé. Téléchargez-le d’abord si vous souhaitez le conserver.</DialogContent><DialogActions><Button disabled={busy} onClick={() => setReloadConfirm(false)}>Annuler</Button><Button disabled={busy} onClick={() => void perform(reload)}>Recharger</Button></DialogActions></Dialog>
    <Snackbar open={Boolean(message)} autoHideDuration={4500} onClose={() => setMessage('')} sx={{ bottom: 'calc(88px + env(safe-area-inset-bottom))' }}><Alert severity="success" variant="filled" onClose={() => setMessage('')}>{message}</Alert></Snackbar>
  </Box>
}
