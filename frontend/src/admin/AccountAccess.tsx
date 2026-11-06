import { createPortal } from 'react-dom'
import { personDisplayName } from '../personDisplayName'
import { accountConfirmation } from './accountIdentity'
import { canManageTeam, isSuperAdmin } from '../rolePolicy'
import { PermanentDeleteDialog } from '../components/PermanentDeleteDialog'
import { useState } from 'react'
import { useQuery } from '@tanstack/react-query'
import { Alert, Box, Button, Chip, Dialog, DialogActions, DialogContent, DialogTitle, Tooltip, Typography } from '@mui/material'
import PersonOffOutlined from '@mui/icons-material/PersonOffOutlined'
import PersonOutlineRounded from '@mui/icons-material/PersonOutlineRounded'
import { adminApi } from './api'
import type { AdminUser, RunAction } from './api'
import DeleteOutlineRounded from '@mui/icons-material/DeleteOutlineRounded'

export function AccountAccess({ token, user, actor, actorRoles, busy, run, confirmChange, requestError, actionsContainer }: {
  actionsContainer?: HTMLElement | null;
  token: string; user: AdminUser; actor: string; actorRoles: string[]; busy: boolean; run: RunAction; confirmChange: (action: () => void) => void; requestError: string;
}) {
  const [confirm, setConfirm] = useState(false)
  const [deleting, setDeleting] = useState<'login' | 'purge' | null>(null)
  const [failed, setFailed] = useState(false)
  const active = user.is_active !== false
  const deleted = !!user.login_deleted_at
  const superAdmin = isSuperAdmin({ role_tags: actorRoles })
  const protectedTarget = isSuperAdmin(user) && !superAdmin
  const confirmationIdentity = accountConfirmation(user)
  const history = useQuery({ queryKey: ['account-activation-history', actor, user.username], queryFn: ({ signal }) => adminApi.activationHistory(token, user.username, signal), retry: false })
  const activationControls = <Box sx={{ display: 'flex', flexWrap: 'wrap', alignItems: 'center', gap: 1,
    '& .MuiButton-root': { whiteSpace: 'nowrap', minHeight: 30, height: 30, px: 1, py: 0, fontSize: 12, fontWeight: 500, borderRadius: 1.5 },
    '& .MuiButton-startIcon': { mr: .5, ml: 0, '& > *:nth-of-type(1)': { fontSize: 16 } } }}>
      <Tooltip title="La désactivation conserve les données ; la suppression définitive est réservée au Super Admin."><Chip size="small" label={deleted ? 'Connexion supprimée · historique conservé' : active ? 'Actif' : 'Désactivé'} variant="outlined" /></Tooltip>
      {!deleted && <Button size="small" variant="outlined" startIcon={active ? <PersonOffOutlined /> : <PersonOutlineRounded />}
        disabled={busy || user.username === actor || protectedTarget} onClick={() => confirmChange(() => { setFailed(false); setConfirm(true) })}>
        {active ? 'Désactiver' : 'Réactiver'}
      </Button>}
  </Box>
  return <Box sx={{ mt: 3 }}>
    {actionsContainer ? createPortal(activationControls, actionsContainer) : activationControls}
    <Box sx={{ display: 'flex', flexWrap: 'wrap', alignItems: 'center', gap: 2 }}>
      {superAdmin && !active && !deleted && user.username !== actor && <Button color="error" variant="outlined" startIcon={<DeleteOutlineRounded />} disabled={busy} onClick={() => confirmChange(() => { setFailed(false); setDeleting('login') })}>Supprimer l’accès</Button>}
      {superAdmin && !active && user.username !== actor && <Button color="error" variant="outlined" startIcon={<DeleteOutlineRounded />} disabled={busy} onClick={() => confirmChange(() => setDeleting('purge'))}>Supprimer définitivement…</Button>}
    </Box>
    {deleted && <Typography variant="body2" color="text.secondary" sx={{ mt: 1 }}>Connexion supprimée ; historique conservé.</Typography>}
    {history.error && <Alert severity="warning" sx={{ mt: 1 }}>Historique des accès indisponible.<Button onClick={() => void history.refetch()}>Réessayer</Button></Alert>}
    {!!history.data?.events.length && <Box sx={{ mt: 2 }}><Typography variant="body2" sx={{ fontWeight: 500 }}>Historique des accès (100 dernières opérations)</Typography>
      {history.data.events.map((event, index) => <Typography key={index} variant="caption" color="text.secondary" sx={{ display: 'block', mt: .5 }}>
        {event.created_at.replace('T', ' ')} · {personDisplayName({ display_name: event.actor_display_name })} · {event.action === 'login_deleted' ? 'Suppression de la connexion · historique conservé' : event.is_active ? 'Réactivation' : 'Désactivation'}
      </Typography>)}
    </Box>}
    <Dialog open={confirm} onClose={() => { if (!busy) setConfirm(false) }}>
      <DialogTitle>{active ? 'Désactiver' : 'Réactiver'} le compte de {personDisplayName(user)} ?</DialogTitle>
      <DialogContent><Alert severity="info">{active ? 'Les sessions seront révoquées. Toutes les données et affectations seront conservées. La personne pourra être réactivée ultérieurement.' : 'La personne pourra se reconnecter avec son mot de passe actuel. Les anciennes sessions resteront invalides.'}</Alert>
        {failed && <Alert severity="error" sx={{ mt: 2 }}>{requestError || 'Opération impossible. Réessayez après actualisation.'}</Alert>}
        {active && canManageTeam(user) && <Typography variant="body2" sx={{ mt: 2 }}>Les personnes rattachées ne sont pas réaffectées automatiquement. Vérifiez leur responsable ; un Admin conserve la possibilité de traiter leurs demandes.</Typography>}
      </DialogContent>
      <DialogActions><Button disabled={busy} onClick={() => setConfirm(false)}>Annuler</Button><Button variant="contained" disabled={busy || protectedTarget || user.username === actor} onClick={async () => {
        setFailed(false)
        const ok = await run(() => adminApi.activation(token, user.username, !active), active ? 'Compte désactivé. Données conservées.' : 'Compte réactivé.')
        if (ok) setConfirm(false)
        else setFailed(true)
      }}>Confirmer</Button></DialogActions>
    </Dialog>
    {deleting && <PermanentDeleteDialog title={deleting === 'login' ? `Supprimer la connexion de ${personDisplayName(user)} ?` : `Supprimer complètement ${personDisplayName(user)} ?`}
      expected={confirmationIdentity.expected} confirmationLabel={confirmationIdentity.label} requireReason={deleting === 'purge'} description={deleting === 'login' ? 'La connexion et le mot de passe seront supprimés. Les données historiques et l’identité seront conservées.' : 'Le compte sera supprimé uniquement s’il ne possède plus de référence métier. Les références existantes bloquent cette opération.'}
      onClose={() => setDeleting(null)} onConfirm={(confirmation, reason) => deleting === 'login' ? adminApi.deleteLogin(token, user.username, confirmation) : adminApi.purge(token, user.username, confirmation, reason)}
      onDeleted={() => { setDeleting(null); void run(async () => {}, deleting === 'login' ? 'Connexion supprimée. Historique conservé.' : 'Compte supprimé définitivement.') }} />}

  </Box>
}
