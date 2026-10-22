import { personDisplayName } from '../personDisplayName'
import { canManageTeam, isAdmin } from '../rolePolicy'
import { useCapabilities } from '../capabilities'
import { useDraftGuard } from '../draftGuard'
import { useRef, useState } from 'react'
import { Accordion, AccordionDetails, AccordionSummary, Alert, Box, Button, Chip, CircularProgress, Dialog, DialogActions, DialogContent, DialogTitle, FormControlLabel, IconButton, Paper, Snackbar, Switch, Tooltip, Typography } from '@mui/material'
import ExpandMoreRounded from '@mui/icons-material/ExpandMoreRounded'
import RefreshRounded from '@mui/icons-material/RefreshRounded'
import { useQuery } from '@tanstack/react-query'
import type { Session } from '../types'
import { dashboardApi } from './api'
import type { Deadline, Interview } from './api'
import { InterviewEditor } from './InterviewEditor'
import { PersonnelDistribution } from './PersonnelDistribution'
import { useQueryClient } from '@tanstack/react-query'
import { WorkspaceToolbar } from '../components/WorkspaceToolbar'

export function DashboardPage({ token, session }: { token: string; session: Session }) {
  const capabilities = useCapabilities()
  const client = useQueryClient()
  const [mailTarget, setMailTarget] = useState<Deadline | 'all' | null>(null)
  const allowed = canManageTeam(session)
  const admin = isAdmin(session)
  const query = useQuery({ queryKey: ['dashboard', session.username], queryFn: ({ signal }) => dashboardApi.get(token, signal), enabled: allowed, retry: false })
  const [editing, setEditing] = useState<Interview | null>(null)
  const [busy, setBusy] = useState(false)
  useDraftGuard(false, busy, 'Tableau de bord')
  const lock = useRef(false)
  const [error, setError] = useState('')
  const [message, setMessage] = useState('')
  async function run(action: () => Promise<unknown>) {
    if (lock.current) return false
    lock.current = true; setBusy(true); setError('')
    try { await action(); await query.refetch(); setMessage('Opération terminée.'); return true }
    catch (e) { setError(e instanceof Error ? e.message : 'Une erreur est survenue.'); return false }
    finally { lock.current = false; setBusy(false) }
  }
  if (!allowed) return <Box className="workspace"><Alert severity="error">Accès réservé aux administrateurs et responsables.</Alert></Box>
  const data = query.data
  return <Box className="workspace" sx={{
    '& .MuiAccordionSummary-root, & .MuiAccordionSummary-root.Mui-expanded': { minHeight: 44 },
    '& .MuiAccordionSummary-content, & .MuiAccordionSummary-content.Mui-expanded': { my: 1 },
  }}>
    <WorkspaceToolbar placement="inline"><Tooltip title="Actualiser le tableau de bord"><span><IconButton aria-label="Actualiser le tableau de bord" disabled={busy || query.isFetching} onClick={() => { void query.refetch(); void client.invalidateQueries({ queryKey: ['organization', session.username] }) }}><RefreshRounded fontSize="small" /></IconButton></span></Tooltip></WorkspaceToolbar>
    {error && <Alert severity="error" sx={{ mb: 2 }}>{error}</Alert>}
    <Box sx={{ display: 'grid', gridTemplateColumns: { xs: 'minmax(0, 1fr)', lg: data ? 'minmax(0, 2fr) minmax(240px, 1fr)' : 'minmax(0, 1fr)' }, gap: 1.5, mb: 1.5, '& > .MuiPaper-root': { mb: 0 } }}>
      <PersonnelDistribution token={token} username={session.username} />
      {data && <Box sx={{ display: 'grid', gridTemplateColumns: { xs: '1fr', sm: 'repeat(3, minmax(0, 1fr))', lg: '1fr' }, gridAutoRows: '1fr', gap: 1.5 }}>{[
        ['Contrats à échéance (10 jours)', data.contracts.length], [`Entretiens à faire (${data.year})`, data.interviews.filter(i => !i.completed).length], ['Entretiens terminés', data.interviews.filter(i => i.completed).length],
      ].map(([label, count]) => <Paper key={label} sx={{ px: 1.75, py: 1.25, minHeight: 64, display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 2 }}><Typography color="text.secondary" variant="body2" sx={{ fontSize: 13 }}>{label}</Typography><Typography sx={{ fontSize: 26, fontWeight: 600, flexShrink: 0, fontVariantNumeric: 'tabular-nums' }}>{count}</Typography></Paper>)}</Box>}
    </Box>
    {query.isPending ? <CircularProgress /> : query.error ? <Alert severity="error">{query.error.message}</Alert> : data && <>
      <Accordion disableGutters><AccordionSummary expandIcon={<ExpandMoreRounded />}><Typography sx={{ fontWeight: 650 }}>Contrats arrivant à échéance</Typography></AccordionSummary><AccordionDetails>
        <Alert severity="info" sx={{ mb: 2 }}>{capabilities.canSendMail ? "Les rappels sont envoyés aux administrateurs et responsables directs après confirmation." : "Les e-mails sont désactivés ou leur configuration ne peut pas être vérifiée."}</Alert>
        {admin && <Box sx={{ display: 'flex', flexWrap: 'wrap', gap: 2, mb: 2 }}><FormControlLabel control={<Switch checked={Boolean(data.contract_reminders_auto_send)} disabled={busy} onChange={(_, value) => void run(() => dashboardApi.setting(token, value))} />} label="Envoi automatique des rappels" /><Button disabled={busy || !capabilities.canSendMail} onClick={() => setMailTarget('all')}>Envoyer les rappels</Button></Box>}
        {!data.contracts.length && <Typography color="text.secondary">Aucun contrat ne se termine dans les 10 prochains jours.</Typography>}
        {data.contracts.map((c, i) => <Paper variant="outlined" key={`${c.employee_username}-${i}`} sx={{ p: 2, mb: 1 }}><Box sx={{ display: 'flex', flexWrap: 'wrap', gap: 2, alignItems: 'center' }}><Typography sx={{ fontWeight: 650, overflowWrap: 'anywhere' }}>{personDisplayName({ display_name: c.employee_name })}</Typography><Chip size="small" variant="outlined" label={`${c.days_remaining} jour(s)`} /><Typography variant="body2">Fin : {c.contract_end_date}</Typography></Box><Typography variant="body2" color="text.secondary" sx={{ mt: 1 }}>Responsable : {personDisplayName({ display_name: c.manager_name })}</Typography>{admin && <Button disabled={busy || !capabilities.canSendMail} size="small" onClick={() => setMailTarget(c)}>Renvoyer le rappel</Button>}</Paper>)}
      </AccordionDetails></Accordion>
      <Accordion disableGutters sx={{ mt: 1.5 }}><AccordionSummary expandIcon={<ExpandMoreRounded />}><Typography sx={{ fontWeight: 650 }}>Entretiens annuels {data.year}</Typography></AccordionSummary><AccordionDetails>
        {!data.interviews.length && <Typography color="text.secondary">Aucun entretien dans votre périmètre.</Typography>}
        {data.interviews.map(item => <Paper variant="outlined" key={item.employee_username} sx={{ p: 2, mb: 1 }}><Box sx={{ display: 'flex', flexWrap: 'wrap', alignItems: 'center', gap: 2 }}><Typography sx={{ flex: 1, overflowWrap: 'anywhere', fontWeight: 650 }}>{personDisplayName({ display_name: item.employee_name })}</Typography><Chip size="small" variant="outlined" label={item.completed ? 'Terminé' : 'À réaliser'} /><Typography variant="body2">Échéance : {item.due_date || '—'}</Typography>{admin && <Button disabled={busy} onClick={() => setEditing(item)}>Modifier</Button>}</Box>
          {item.notes && <Typography variant="body2" color="text.secondary" sx={{ mt: 1, whiteSpace: 'pre-wrap', overflowWrap: 'anywhere' }}>{item.notes}</Typography>}
          {item.completed && <Typography variant="caption" color="text.secondary">Terminé par {personDisplayName({ display_name: item.completed_by_name })} · {item.completed_at || '—'}</Typography>}
        </Paper>)}
      </AccordionDetails></Accordion>
    </>}
    {editing && <InterviewEditor key={`${editing.employee_username}-${editing.interview_year}`} item={editing} busy={busy} onClose={() => setEditing(null)} save={values => run(() => dashboardApi.interview(token, editing, values))} />}
    <Dialog open={mailTarget !== null} onClose={() => { if (!busy) setMailTarget(null) }}><DialogTitle>Confirmer l’envoi des rappels</DialogTitle><DialogContent>De vrais e-mails seront envoyés aux administrateurs et responsables directs pour {mailTarget === 'all' ? 'les contrats actuellement à échéance' : personDisplayName({ display_name: mailTarget?.employee_name })}. Un renvoi peut répéter un e-mail déjà reçu.</DialogContent><DialogActions><Button disabled={busy} onClick={() => setMailTarget(null)}>Annuler</Button><Button disabled={busy || !capabilities.canSendMail} onClick={() => { if (!mailTarget || !capabilities.canSendMail) return; const target = mailTarget; void run(() => target === 'all' ? dashboardApi.send(token) : dashboardApi.resend(token, target)).then(ok => { if (ok) setMailTarget(null) }) }}>Envoyer les e-mails</Button></DialogActions></Dialog>
    <Snackbar open={!!message} autoHideDuration={5000} onClose={() => setMessage('')}><Alert severity="success" variant="filled" onClose={() => setMessage('')}>{message}</Alert></Snackbar>
  </Box>
}
