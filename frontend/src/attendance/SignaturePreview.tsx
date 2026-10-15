import { useEffect, useRef, useState } from 'react'
import { Alert, Button, Checkbox, CircularProgress, Dialog, DialogActions, DialogContent, DialogTitle, FormControlLabel, LinearProgress, Paper, Typography } from '@mui/material'
import { useQuery } from '@tanstack/react-query'
import { request } from '../api'
interface PreviewItem { id: string; name: string; employee_email: string; manager_name: string; manager_email: string; error: string; fingerprint?: string }
interface SigningAction { url: string; party: 'employee' | 'manager' | 'unknown' }
interface SendResult { status: string; document_id: string; message: string; signing_links?: string[]; signing_actions?: SigningAction[]; timings?: { total_seconds: number } }
function signerActions(result: { document_id?: string; links?: string[]; actions?: SigningAction[] }) {
  const source = result.actions?.length ? result.actions : (result.links || []).map(url => ({ url, party: 'unknown' as const }))
  return source.filter(({ url }) => {
    try {
      const parsed = new URL(url)
      return parsed.protocol === 'https:' && (parsed.hostname === 'clawshow.ai' || parsed.hostname.endsWith('.clawshow.ai')) && !parsed.username && parsed.pathname.replace(/\/$/, '') === `/esign/${result.document_id}`
    } catch { return false }
  })
}
export function SignaturePreview({ token, ids, month, year, onClose }: { token: string; ids: string[]; month: number; year: number; onClose: () => void }) {
  const preview = useQuery({ queryKey: ['signature-preview', ids, month, year], queryFn: ({ signal }) => request<{ items: PreviewItem[]; reason: string; can_send: boolean; mode: string }>('/attendance/signature-preview', token,
    { method: 'POST', signal, body: JSON.stringify({ ids, month, year }) }), retry: false, staleTime: 0, gcTime: 0, refetchOnWindowFocus: false })
  const [confirmed, setConfirmed] = useState(false)
  const [busy, setBusy] = useState(false)
  const lock = useRef(false)
  const [outcomes, setOutcomes] = useState<Record<string, { ok: boolean; message: string; document_id?: string; links?: string[]; actions?: SigningAction[]; seconds?: number }>>({})
  const [seconds, setSeconds] = useState(0)
  useEffect(() => {
    if (!busy) return
    const timer = window.setInterval(() => setSeconds(value => value + 1), 1000)
    return () => window.clearInterval(timer)
  }, [busy])
  const attempts = useRef<Record<string, string>>({})
  const [processed, setProcessed] = useState(0)
  const [started, setStarted] = useState(false)
  const live = ['production', 'production-test'].includes(preview.data?.mode || '')
  async function send() {
    if (lock.current || !confirmed || !preview.data?.can_send) return
    lock.current = true; setBusy(true); setStarted(true); setSeconds(0)
    try {
      for (const item of preview.data.items) {
        if (item.error || !item.fingerprint) {
          setOutcomes(o => ({ ...o, [item.id]: { ok: false, message: item.error || 'Fiche invalide.' } }))
        } else {
          const request_id = attempts.current[item.id] ||= crypto.randomUUID()
          try {
            const result = await request<SendResult>('/attendance/signature-send', token, { method: 'POST', body: JSON.stringify({ id: item.id, month, year, fingerprint: item.fingerprint, request_id }) }, { timeout: 240000 })
            setOutcomes(o => ({ ...o, [item.id]: { ok: true, message: result.message, document_id: result.document_id, links: result.signing_links || [], actions: result.signing_actions, seconds: result.timings?.total_seconds } }))
          } catch (e) {
            setOutcomes(o => ({ ...o, [item.id]: { ok: false, message: `${e instanceof Error ? e.message : 'Envoi interrompu.'} Vérifiez ClawShow avant tout nouvel envoi.` } }))
          }
        }
        setProcessed(value => value + 1)
      }
    } finally { lock.current = false; setBusy(false) }
  }
  return <Dialog open fullWidth maxWidth="md" onClose={() => { if (!busy) onClose() }}><DialogTitle>Préparer les signatures · {month.toString().padStart(2, '0')}/{year}</DialogTitle><DialogContent>
    <Alert severity="info" sx={{ mb: 2 }}>Le salarié signe en premier, puis son responsable direct enregistré dans Administration. Le téléchargement et la signature sont deux opérations distinctes.</Alert>
    {preview.isPending ? <CircularProgress /> : preview.error ? <Alert severity="error">{preview.error.message}</Alert> : preview.data && <>
      <Alert severity={live || !preview.data.can_send ? 'warning' : 'info'} sx={{ mb: 2 }}>{preview.data.reason}</Alert>
      {busy && <><LinearProgress /><Typography variant="body2" sx={{ my: 1 }}>Préparation du PDF et création de la demande ClawShow… {seconds} s · {processed}/{preview.data.items.length} fiche(s) traitée(s). Gardez cette fenêtre ouverte.</Typography></>}
      {started && !busy && <Alert severity="info" sx={{ mb: 2 }}>Traitement terminé : {Object.values(outcomes).filter(o => o.ok).length} demande(s) créée(s), {Object.values(outcomes).filter(o => !o.ok).length} échec(s). Les signatures restent à effectuer.</Alert>}
      {preview.data.items.map(item => <Paper key={item.id} variant="outlined" sx={{ p: 2, mb: 1, overflowWrap: 'anywhere' }}><Typography sx={{ fontWeight: 650 }}>{item.name || `Fiche ${item.id}`}</Typography>
        <Typography variant="body2">1. Salarié : {item.employee_email || 'E-mail manquant'}</Typography><Typography variant="body2">2. Responsable : {item.manager_name || '—'} · {item.manager_email || '—'}</Typography>
        {item.error && <Alert severity="error" sx={{ mt: 1 }}>{item.error}</Alert>}
        {outcomes[item.id] && <Alert severity={outcomes[item.id].ok ? 'success' : 'error'} sx={{ mt: 1 }}>{outcomes[item.id].ok ? 'Demande créée — signatures en attente. ' : ''}{outcomes[item.id].message}{outcomes[item.id].document_id && ` Référence : ${outcomes[item.id].document_id}`}</Alert>}
        {outcomes[item.id] && signerActions(outcomes[item.id]).map((action, index) => <Button key={action.url} component="a" href={action.url} target="_blank" rel="noopener noreferrer" sx={{ mt: 1 }}>{action.party === 'employee' ? 'Signature du salarié' : action.party === 'manager' ? 'Signature du responsable' : `Ouvrir la page de signature ${index + 1}`}</Button>)}
        {outcomes[item.id]?.seconds !== undefined && <Typography variant="caption" sx={{ display: 'block', mt: 1 }}>Demande préparée en {outcomes[item.id].seconds} s.</Typography>}
        {outcomes[item.id]?.ok && live && <Typography variant="body2" sx={{ mt: 1 }}>Consultez l’e-mail de signature envoyé au salarié. Le responsable recevra son invitation après la première signature.</Typography>}
        {outcomes[item.id]?.ok && !live && !signerActions(outcomes[item.id]).length && <Typography variant="body2" sx={{ mt: 1 }}>ClawShow n’a pas renvoyé de page de signature identifiable. Conservez la référence pour vérifier la demande.</Typography>}
      </Paper>)}
      <FormControlLabel control={<Checkbox checked={confirmed} disabled={busy || started || !preview.data.can_send} onChange={e => setConfirmed(e.target.checked)} />} label={live ? 'Je confirme le document et l’envoi réel aux destinataires affichés. Cette demande utilise le quota de signatures Production.' : 'Je confirme les destinataires. Une nouvelle demande sera créée pour chaque fiche valide ; vérifiez les envois précédents pour éviter les doublons.'} />
    </>}
  </DialogContent><DialogActions><Button disabled={busy} onClick={onClose}>Fermer</Button><Button variant="contained" disabled={busy || started || !confirmed || !preview.data?.can_send || !preview.data.items.some(item => !item.error)} onClick={() => void send()}>{busy ? 'Création en cours…' : started ? 'Traitement terminé' : preview.data?.can_send ? (live ? 'Envoyer les e-mails de signature' : 'Envoyer en sandbox') : 'Envoi indisponible (configuration requise)'}</Button></DialogActions></Dialog>
}
