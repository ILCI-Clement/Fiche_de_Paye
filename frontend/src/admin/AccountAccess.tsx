import { useState } from 'react'
import { useQuery } from '@tanstack/react-query'
import { Alert, Box, Button, Chip, Dialog, DialogActions, DialogContent, DialogTitle, TextField, Typography } from '@mui/material'
import PersonOffOutlined from '@mui/icons-material/PersonOffOutlined'
import PersonOutlineRounded from '@mui/icons-material/PersonOutlineRounded'
import { adminApi } from './api'
import type { AdminUser, RunAction } from './api'
import { DeleteAction } from '../components/DeleteAction'

export function AccountAccess({ token, user, actor, busy, run, confirmChange, requestError }: {
  token: string; user: AdminUser; actor: string; busy: boolean; run: RunAction; confirmChange: (action: () => void) => void; requestError: string;
}) {
  const [confirm, setConfirm] = useState(false)
  const [deleting, setDeleting] = useState(false)
  const [confirmation, setConfirmation] = useState('')
  const [failed, setFailed] = useState(false)
  const active = user.is_active !== false
  const deleted = !!user.login_deleted_at
  const history = useQuery({ queryKey: ['account-activation-history', actor, user.username], queryFn: ({ signal }) => adminApi.activationHistory(token, user.username, signal), retry: false })
  return <Box sx={{ mt: 3 }}>
    <Box sx={{ display: 'flex', flexWrap: 'wrap', alignItems: 'center', gap: 2 }}>
      <Chip size="small" label={deleted ? 'Connexion supprimée · historique conservé' : active ? 'Compte actif' : 'Compte désactivé'} variant="outlined" />
      {!deleted && <Button variant="outlined" startIcon={active ? <PersonOffOutlined /> : <PersonOutlineRounded />}
        disabled={busy || user.username === actor} onClick={() => confirmChange(() => { setFailed(false); setConfirm(true) })}>
        {active ? 'Désactiver le compte' : 'Réactiver le compte'}
      </Button>}
      {!active && !deleted && user.username !== actor && <DeleteAction label="Supprimer la connexion et conserver l’historique" disabled={busy} onClick={() => confirmChange(() => { setFailed(false); setConfirmation(''); setDeleting(true) })} />}
    </Box>
    <Typography variant="body2" color="text.secondary" sx={{ mt: 1 }}>{deleted ? 'L’accès est définitivement supprimé. L’identité historique et les données restent consultables ; ce compte ne peut plus être réactivé ni recréé.' : 'Désactivez d’abord le compte. Vous pourrez ensuite supprimer sa connexion en conservant les présences, contrats, documents et historiques.'}</Typography>
    {history.error && <Alert severity="warning" sx={{ mt: 1 }}>Historique des accès indisponible.<Button onClick={() => void history.refetch()}>Réessayer</Button></Alert>}
    {!!history.data?.events.length && <Box sx={{ mt: 2 }}><Typography variant="body2" sx={{ fontWeight: 500 }}>Historique des accès (100 dernières opérations)</Typography>
      {history.data.events.map((event, index) => <Typography key={index} variant="caption" color="text.secondary" sx={{ display: 'block', mt: .5 }}>
        {event.created_at.replace('T', ' ')} · {event.actor_username} · {event.action === 'login_deleted' ? 'Suppression de la connexion · historique conservé' : event.is_active ? 'Réactivation' : 'Désactivation'}
      </Typography>)}
    </Box>}
    <Dialog open={confirm} onClose={() => { if (!busy) setConfirm(false) }}>
      <DialogTitle>{active ? 'Désactiver' : 'Réactiver'} le compte {user.username} ?</DialogTitle>
      <DialogContent><Alert severity="info">{active ? 'Les sessions seront révoquées. Toutes les données et affectations seront conservées. La personne pourra être réactivée ultérieurement.' : 'La personne pourra se reconnecter avec son mot de passe actuel. Les anciennes sessions resteront invalides.'}</Alert>
        {failed && <Alert severity="error" sx={{ mt: 2 }}>{requestError || 'Opération impossible. Réessayez après actualisation.'}</Alert>}
        {active && user.role_tags.some(role => ['Admin', 'Responsable'].includes(role)) && <Typography variant="body2" sx={{ mt: 2 }}>Les personnes rattachées ne sont pas réaffectées automatiquement. Vérifiez leur responsable ; un Admin conserve la possibilité de traiter leurs demandes.</Typography>}
      </DialogContent>
      <DialogActions><Button disabled={busy} onClick={() => setConfirm(false)}>Annuler</Button><Button variant="contained" disabled={busy} onClick={async () => {
        setFailed(false)
        const ok = await run(() => adminApi.activation(token, user.username, !active), active ? 'Compte désactivé. Données conservées.' : 'Compte réactivé.')
        if (ok) setConfirm(false)
        else setFailed(true)
      }}>Confirmer</Button></DialogActions>
    </Dialog>
    <Dialog open={deleting} onClose={() => { if (!busy) setDeleting(false) }}>
      <DialogTitle>Supprimer la connexion de {user.username} ?</DialogTitle>
      <DialogContent>
        <Alert severity="warning" sx={{ mb: 2 }}>La connexion et le mot de passe seront supprimés définitivement. Le compte ne pourra plus être réactivé. Les présences, contrats, fichiers (y compris ceux déposés pour d’autres personnes) et historiques seront conservés.</Alert>
        <Typography variant="body2" color="text.secondary" sx={{ mb: 2 }}>Le nom de compte et l’adresse e-mail restent réservés pour protéger le rattachement des archives.</Typography>
        <TextField fullWidth label={`Saisissez ${user.username} pour confirmer`} value={confirmation} disabled={busy} onChange={event => setConfirmation(event.target.value)} autoComplete="off" />
        {failed && <Alert severity="error" sx={{ mt: 2 }}>{requestError || 'Suppression impossible. Réessayez après actualisation.'}</Alert>}
      </DialogContent>
      <DialogActions><Button disabled={busy} onClick={() => setDeleting(false)}>Annuler</Button><Button color="error" variant="contained" disabled={busy || confirmation !== user.username} onClick={async () => {
        setFailed(false)
        const ok = await run(() => adminApi.deleteLogin(token, user.username, confirmation), 'Connexion supprimée. Historique conservé.')
        if (ok) setDeleting(false)
        else setFailed(true)
      }}>Supprimer la connexion</Button></DialogActions>
    </Dialog>
  </Box>
}
