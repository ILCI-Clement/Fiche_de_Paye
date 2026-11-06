import { personDisplayName } from '../personDisplayName'
import { useDraftGuard } from '../draftGuard'
import { useState } from 'react'
import { Alert, Box, Button, CircularProgress, Dialog, DialogActions, DialogContent, DialogTitle, Divider, Stack, TextField, Typography } from '@mui/material'
import { useMutation, useQuery } from '@tanstack/react-query'
import { api } from '../api'
import { periodLabel, slotLabel, statusLabels, timeLabel } from '../leaveFormat'
import { StatusChip } from './StatusChip'
import { LeaveAttendanceCheck } from './LeaveAttendanceCheck'
import { useWorkNavigation } from '../workspaceNavigation'

type Action = 'approved' | 'rejected' | 'withdrawn' | 'cancel_requested' | 'cancel_approved' | 'cancel_rejected'
const cancellationLabels = { cancel_requested: 'Annulation demandée', cancel_approved: 'Annulation approuvée', cancel_rejected: 'Annulation refusée' }

export function LeaveDetail({ token, id, onClose, onChanged }: {
  token: string; id: number; onClose: () => void; onChanged: (message: string) => void
}) {
  const navigation = useWorkNavigation()
  const [accepted, setAccepted] = useState(false)
  const [comment, setComment] = useState('')
  const [confirm, setConfirm] = useState<Action | null>(null)
  const detail = useQuery({ queryKey: ['leave-detail', id], queryFn: ({ signal }) => api.detail(token, id, signal), refetchInterval: 30000 })
  const operation = confirm === 'approved' ? 'approve' : confirm === 'cancel_approved' ? 'cancel' : undefined
  const preview = useQuery({ queryKey: ['leave-attendance-preview', id, operation],
    queryFn: ({ signal }) => api.leaveAttendancePreview(token, id, operation!, signal), enabled: !!operation, retry: false })
  const choose = (action: Action) => { setAccepted(false); setConfirm(action) }
  const mutation = useMutation({
    mutationFn: (action: Action) => {
      if (action === 'cancel_requested') return api.requestCancellation(token, id)
      if (action === 'cancel_approved' || action === 'cancel_rejected') return api.decideCancellation(token, id, action === 'cancel_approved' ? 'approved' : 'rejected', comment, preview.data?.fingerprint)
      return action === 'withdrawn' ? api.withdraw(token, id) : api.decide(token, id, action, comment, action === 'approved' ? preview.data?.fingerprint : undefined)
    },
    onSuccess: (_, action) => onChanged(action in cancellationLabels ? cancellationLabels[action as keyof typeof cancellationLabels] : 'Demande mise à jour.'),
    onError: () => { setConfirm(null); void detail.refetch() },
  })
  const close = useDraftGuard(!!comment, mutation.isPending, 'Décision de congé')
  const item = detail.data?.request
  return <><Dialog open onClose={mutation.isPending ? undefined : () => close(onClose)} fullWidth maxWidth="sm">
    <DialogTitle sx={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: 1 }}>
      Demande #{id}{item && <StatusChip status={item.status} cancellationPending={item.cancellation_status === 'pending'} />}
    </DialogTitle>
    <DialogContent>
      {detail.isPending && <CircularProgress size={24} />}
      {detail.error && <Alert severity="error">{detail.error.message}</Alert>}
      {mutation.error && <Alert severity="error" sx={{ mb: 2 }}>{mutation.error.message}</Alert>}
      {item && <>
        <Typography variant="h6">{personDisplayName({ display_name: item.employee_name })}</Typography>
        <Typography sx={{ mt: 2, fontWeight: 650 }}>{periodLabel(item)}</Typography>
        <Typography color="text.secondary" variant="body2">{slotLabel(item)}</Typography>
        <Box sx={{ p: 2, bgcolor: 'var(--surface-soft)', borderRadius: 2, my: 2 }}>
          <Typography variant="caption" color="text.secondary">Note de l’employé</Typography>
          <Typography sx={{ whiteSpace: 'pre-wrap', overflowWrap: 'anywhere' }}>{item.note || 'Aucune note ajoutée.'}</Typography>
        </Box>
        <Typography variant="body2" color="text.secondary">Responsable actuel : <strong>{personDisplayName({ display_name: item.approver_name })}</strong></Typography>
        <Divider sx={{ my: 3 }} />
        <Typography sx={{ fontWeight: 700, mb: 2 }}>Historique des opérations</Typography>
        <Stack spacing={2}>
          {detail.data?.history.map(event => <Box key={event.id} sx={{ pl: 2, borderLeft: '2px solid', borderColor: 'divider' }}>
            <Typography variant="body2" sx={{ fontWeight: 650 }}>{event.action === 'attendance_synced' ? 'Présences synchronisées' : event.action in cancellationLabels ? cancellationLabels[event.action as keyof typeof cancellationLabels] : event.action === 'submitted' ? 'Demande envoyée' : statusLabels[event.new_status]} · {personDisplayName({ display_name: event.actor_name })}</Typography>
            <Typography variant="caption" color="text.secondary">{timeLabel(event.created_at)}</Typography>
            {event.comment && <Typography variant="body2" sx={{ mt: 0.5, whiteSpace: 'pre-wrap', overflowWrap: 'anywhere' }}>{event.comment}</Typography>}
          </Box>)}
        </Stack>
        {(item.can_review || item.can_review_cancellation) && <TextField sx={{ mt: 3 }} fullWidth multiline minRows={2} label="Commentaire de décision (facultatif)"
          value={comment} disabled={mutation.isPending} onChange={event => setComment(event.target.value)} slotProps={{ htmlInput: { maxLength: 2000 } }} />}

      </>}
    </DialogContent>
    <DialogActions sx={{ flexWrap: 'wrap', gap: 1, px: 3, pb: 3 }}>
      <Button color="inherit" onClick={() => close(onClose)} disabled={mutation.isPending}>Fermer</Button>
      {item?.can_withdraw && <Button color="inherit" variant="outlined" disabled={mutation.isPending} onClick={() => choose('withdrawn')}>Annuler ma demande</Button>}
      {item?.can_request_cancellation && <Button variant="outlined" disabled={mutation.isPending} onClick={() => choose('cancel_requested')}>Demander l’annulation</Button>}
      {item?.can_review_cancellation && <>
        <Button variant="outlined" disabled={mutation.isPending} onClick={() => choose('cancel_rejected')}>Refuser l’annulation</Button>
        <Button variant="contained" disabled={mutation.isPending} onClick={() => choose('cancel_approved')}>Approuver l’annulation</Button>
      </>}
      {item?.can_review && <>
        <Button color="error" variant="outlined" disabled={mutation.isPending} onClick={() => choose('rejected')}>Refuser</Button>
        <Button variant="contained" disabled={mutation.isPending} onClick={() => choose('approved')}>Approuver</Button>
      </>}
    </DialogActions>
  </Dialog>
    <Dialog open={confirm !== null} onClose={mutation.isPending ? undefined : () => setConfirm(null)} fullWidth maxWidth="xs"
      aria-labelledby="leave-confirm-title" aria-describedby="leave-confirm-message">
      <DialogTitle id="leave-confirm-title">Confirmer l’action</DialogTitle>
      <DialogContent><Typography id="leave-confirm-message">{confirm === 'cancel_requested' ? 'Demander l’annulation ? Le congé reste approuvé jusqu’à validation.' : confirm === 'cancel_approved' ? 'Approuver l’annulation de ce congé ?' : confirm === 'cancel_rejected' ? 'Refuser l’annulation et conserver ce congé ?' : confirm === 'withdrawn' ? 'Annuler cette demande en attente ?' : `${confirm === 'approved' ? 'Approuver' : 'Refuser'} cette demande ?`}</Typography>
        {operation && <>
          {preview.isFetching && <CircularProgress size={20} sx={{ mt: 2 }} />}
          {preview.error && <Alert severity="error" sx={{ mt: 2 }}>{preview.error.message}<Button onClick={() => void preview.refetch()}>Réessayer</Button></Alert>}
          {!preview.isFetching && preview.data && <LeaveAttendanceCheck data={preview.data} accepted={accepted} onAccepted={setAccepted} />}
          {preview.data?.issues.some(issue => issue.blocking) && navigation && item && <Button sx={{ mt: 1 }}
            onClick={() => close(() => { onClose(); navigation.navigate({ page: 'attendance', username: item.employee_username || undefined,
              month: Number((preview.data?.issues[0]?.date || item.start_date).slice(5, 7)),
              year: Number((preview.data?.issues[0]?.date || item.start_date).slice(0, 4)) }) })}>Ouvrir les fiches de temps</Button>}
        </>}
      </DialogContent>
      <DialogActions sx={{ px: 3, pb: 2, gap: 1 }}>
        <Button autoFocus color="inherit" disabled={mutation.isPending} onClick={() => setConfirm(null)}>Revenir</Button>
        <Button variant="contained" disabled={mutation.isPending || confirm === null || (!!operation && (preview.isFetching || !preview.data?.can_apply || (preview.data.needs_confirmation && !accepted)))}
          onClick={() => { if (confirm) mutation.mutate(confirm) }}>
          {mutation.isPending ? 'En cours…' : 'Confirmer'}
        </Button>
      </DialogActions>
    </Dialog>
  </>
}
