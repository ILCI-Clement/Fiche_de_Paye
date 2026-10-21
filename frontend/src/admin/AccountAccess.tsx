import { useState } from 'react'
import { useQuery } from '@tanstack/react-query'
import { Alert, Box, Button, Chip, Dialog, DialogActions, DialogContent, DialogTitle, Typography } from '@mui/material'
import PersonOffOutlined from '@mui/icons-material/PersonOffOutlined'
import PersonOutlineRounded from '@mui/icons-material/PersonOutlineRounded'
import { adminApi } from './api'
import type { AdminUser, RunAction } from './api'

export function AccountAccess({ token, user, actor, busy, run, confirmChange }: {
  token: string; user: AdminUser; actor: string; busy: boolean; run: RunAction; confirmChange: (action: () => void) => void;
}) {
  const [confirm, setConfirm] = useState(false)
  const active = user.is_active !== false
  const history = useQuery({ queryKey: ['account-activation-history', actor, user.username], queryFn: ({ signal }) => adminApi.activationHistory(token, user.username, signal), retry: false })
  return <Box sx={{ mt: 3 }}>
    <Box sx={{ display: 'flex', flexWrap: 'wrap', alignItems: 'center', gap: 2 }}>
      <Chip size="small" label={active ? 'Compte actif' : 'Compte désactivé'} variant="outlined" />
      <Button variant="outlined" startIcon={active ? <PersonOffOutlined /> : <PersonOutlineRounded />}
        disabled={busy || user.username === actor} onClick={() => confirmChange(() => setConfirm(true))}>
        {active ? 'Désactiver le compte' : 'Réactiver le compte'}
      </Button>
    </Box>
    <Typography variant="body2" color="text.secondary" sx={{ mt: 1 }}>La désactivation bloque l’accès et conserve les présences, contrats, documents et historiques.</Typography>
    {history.error && <Alert severity="warning" sx={{ mt: 1 }}>Historique des accès indisponible.<Button onClick={() => void history.refetch()}>Réessayer</Button></Alert>}
    {!!history.data?.events.length && <Box sx={{ mt: 2 }}><Typography variant="body2" sx={{ fontWeight: 500 }}>Historique des accès (100 dernières opérations)</Typography>
      {history.data.events.map((event, index) => <Typography key={index} variant="caption" color="text.secondary" sx={{ display: 'block', mt: .5 }}>
        {event.created_at.replace('T', ' ')} · {event.actor_username} · {event.is_active ? 'Réactivation' : 'Désactivation'}
      </Typography>)}
    </Box>}
    <Dialog open={confirm} onClose={() => { if (!busy) setConfirm(false) }}>
      <DialogTitle>{active ? 'Désactiver' : 'Réactiver'} le compte {user.username} ?</DialogTitle>
      <DialogContent><Alert severity="info">{active ? 'Les sessions seront révoquées. Toutes les données et affectations seront conservées. La personne pourra être réactivée ultérieurement.' : 'La personne pourra se reconnecter avec son mot de passe actuel. Les anciennes sessions resteront invalides.'}</Alert>
        {active && user.role_tags.some(role => ['Admin', 'Responsable'].includes(role)) && <Typography variant="body2" sx={{ mt: 2 }}>Les personnes rattachées ne sont pas réaffectées automatiquement. Vérifiez leur responsable ; un Admin conserve la possibilité de traiter leurs demandes.</Typography>}
      </DialogContent>
      <DialogActions><Button disabled={busy} onClick={() => setConfirm(false)}>Annuler</Button><Button variant="contained" disabled={busy} onClick={async () => {
        const ok = await run(() => adminApi.activation(token, user.username, !active), active ? 'Compte désactivé. Données conservées.' : 'Compte réactivé.')
        if (ok) setConfirm(false)
      }}>Confirmer</Button></DialogActions>
    </Dialog>
  </Box>
}
