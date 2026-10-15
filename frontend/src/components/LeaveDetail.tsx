import { useDraftGuard } from '../draftGuard'
import { useState } from 'react'
import { Alert, Box, Button, CircularProgress, Dialog, DialogActions, DialogContent, DialogTitle, Divider, Stack, TextField, Typography } from '@mui/material'
import { useMutation, useQuery } from '@tanstack/react-query'
import { api } from '../api'
import { periodLabel, slotLabel, statusLabels, timeLabel } from '../leaveFormat'
import { StatusChip } from './StatusChip'

export function LeaveDetail({ token, id, onClose, onChanged }: {
  token: string; id: number; onClose: () => void; onChanged: (message: string) => void
}) {
  const [comment, setComment] = useState('')
  const [confirm, setConfirm] = useState<'approved' | 'rejected' | 'withdrawn' | null>(null)
  const detail = useQuery({ queryKey: ['leave-detail', id], queryFn: ({ signal }) => api.detail(token, id, signal), refetchInterval: 30000 })
  const mutation = useMutation({
    mutationFn: (action: 'approved' | 'rejected' | 'withdrawn') => action === 'withdrawn'
      ? api.withdraw(token, id) : api.decide(token, id, action, comment),
    onSuccess: item => onChanged(item.status === 'withdrawn' ? 'Demande retirée.' : `Demande ${item.status === 'approved' ? 'approuvée' : 'refusée'}.`),
    onError: () => { setConfirm(null); void detail.refetch() },
  })
  const close = useDraftGuard(!!comment, mutation.isPending, 'Décision de congé')
  const item = detail.data?.request
  return <Dialog open onClose={mutation.isPending ? undefined : () => close(onClose)} fullWidth maxWidth="sm">
    <DialogTitle sx={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: 1 }}>
      Demande #{id}{item && <StatusChip status={item.status} />}
    </DialogTitle>
    <DialogContent>
      {detail.isPending && <CircularProgress size={24} />}
      {detail.error && <Alert severity="error">{detail.error.message}</Alert>}
      {mutation.error && <Alert severity="error" sx={{ mb: 2 }}>{mutation.error.message}</Alert>}
      {item && <>
        <Typography variant="h6">{item.employee_name}</Typography>
        <Typography sx={{ mt: 2, fontWeight: 650 }}>{periodLabel(item)}</Typography>
        <Typography color="text.secondary" variant="body2">{slotLabel(item)}</Typography>
        <Box sx={{ p: 2, bgcolor: 'var(--surface-soft)', borderRadius: 2, my: 2 }}>
          <Typography variant="caption" color="text.secondary">Note de l’employé</Typography>
          <Typography sx={{ whiteSpace: 'pre-wrap', overflowWrap: 'anywhere' }}>{item.note || 'Aucune note ajoutée.'}</Typography>
        </Box>
        <Typography variant="body2" color="text.secondary">Responsable actuel : <strong>{item.approver_name}</strong></Typography>
        <Divider sx={{ my: 3 }} />
        <Typography sx={{ fontWeight: 700, mb: 2 }}>Historique des opérations</Typography>
        <Stack spacing={2}>
          {detail.data?.history.map(event => <Box key={event.id} sx={{ pl: 2, borderLeft: '2px solid', borderColor: 'divider' }}>
            <Typography variant="body2" sx={{ fontWeight: 650 }}>{event.action === 'submitted' ? 'Demande envoyée' : statusLabels[event.new_status]} · {event.actor_name}</Typography>
            <Typography variant="caption" color="text.secondary">{timeLabel(event.created_at)}</Typography>
            {event.comment && <Typography variant="body2" sx={{ mt: 0.5, whiteSpace: 'pre-wrap', overflowWrap: 'anywhere' }}>{event.comment}</Typography>}
          </Box>)}
        </Stack>
        {item.can_review && <TextField sx={{ mt: 3 }} fullWidth multiline minRows={2} label="Commentaire de décision (facultatif)"
          value={comment} disabled={mutation.isPending} onChange={event => setComment(event.target.value)} slotProps={{ htmlInput: { maxLength: 2000 } }} />}
        {confirm && <Alert severity={confirm === 'approved' ? 'success' : 'warning'} sx={{ mt: 2 }}
          action={<Button color="inherit" disabled={mutation.isPending} onClick={() => mutation.mutate(confirm)}>Confirmer</Button>}>
          {confirm === 'withdrawn' ? 'Retirer cette demande ? Elle disparaîtra de la liste.' : `${confirm === 'approved' ? 'Approuver' : 'Refuser'} cette demande ?`}
        </Alert>}
      </>}
    </DialogContent>
    <DialogActions sx={{ flexWrap: 'wrap', gap: 1, px: 3, pb: 3 }}>
      <Button color="inherit" onClick={() => close(onClose)} disabled={mutation.isPending}>Fermer</Button>
      {item?.can_withdraw && <Button color="inherit" variant="outlined" disabled={mutation.isPending} onClick={() => setConfirm('withdrawn')}>Retirer ma demande</Button>}
      {item?.can_review && <>
        <Button color="error" variant="outlined" disabled={mutation.isPending} onClick={() => setConfirm('rejected')}>Refuser</Button>
        <Button variant="contained" disabled={mutation.isPending} onClick={() => setConfirm('approved')}>Approuver</Button>
      </>}
    </DialogActions>
  </Dialog>
}
