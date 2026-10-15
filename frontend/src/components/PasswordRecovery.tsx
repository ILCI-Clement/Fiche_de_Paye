import { useCapabilities } from '../capabilities'
import { useState } from 'react'
import { Alert, Button, Dialog, DialogActions, DialogContent, DialogTitle, Stack, TextField } from '@mui/material'
import { request } from '../api'

export function PasswordRecovery({ onClose, initialToken = '' }: { onClose: () => void; initialToken?: string }) {
  const [mode, setMode] = useState<'request' | 'reset'>(initialToken ? 'reset' : 'request')
  const [email, setEmail] = useState('')
  const [token, setToken] = useState(initialToken)
  const [password, setPassword] = useState('')
  const [confirmation, setConfirmation] = useState('')
  const [busy, setBusy] = useState(false)
  const [message, setMessage] = useState('')
  const [error, setError] = useState('')
  const capabilities = useCapabilities()
  const local = !capabilities.canSendMail
  async function submit() {
    if (busy || (local && mode === 'request')) return
    if (mode === 'reset' && password !== confirmation) { setError('Les mots de passe ne correspondent pas.'); return }
    setBusy(true); setError(''); setMessage('')
    try {
      if (mode === 'request') {
        await request('/forgot-password', null, { method: 'POST', body: JSON.stringify({ email }) })
        setMessage('Si cette adresse correspond à un compte, un e-mail de réinitialisation a été envoyé.')
      } else {
        await request('/reset-password', null, { method: 'POST', body: JSON.stringify({ token: token.trim(), new_password: password }) })
        setToken(''); setPassword(''); setConfirmation(''); setMessage('Mot de passe réinitialisé. Vous pouvez vous connecter.')
      }
    } catch (e) { setError(e instanceof Error ? e.message : 'Une erreur est survenue.') }
    finally { setBusy(false) }
  }
  return <Dialog open onClose={() => { if (!busy) onClose() }} fullWidth maxWidth="sm"><DialogTitle>Réinitialiser le mot de passe</DialogTitle>
    <form onSubmit={e => { e.preventDefault(); void submit() }}><DialogContent><Stack spacing={2}>
      {local && <Alert severity="info">L’envoi des e-mails est désactivé ou sa configuration ne peut pas être vérifiée. Aucun lien de réinitialisation ne sera envoyé. Un jeton obtenu dans un autre environnement ne fonctionne pas ici.</Alert>}
      {mode === 'request' ? <TextField label="Adresse e-mail" type="email" required autoComplete="email" value={email} onChange={e => setEmail(e.target.value)} disabled={busy} helperText={local ? 'Vous pouvez saisir votre adresse, mais aucun lien ne peut être envoyé dans cet environnement.' : undefined} /> : <>
        <TextField label="Jeton de réinitialisation" type="password" required value={token} onChange={e => setToken(e.target.value)} disabled={busy} />
        <TextField label="Nouveau mot de passe" type="password" autoComplete="new-password" required value={password} onChange={e => setPassword(e.target.value)} disabled={busy} slotProps={{ htmlInput: { minLength: 8 } }} />
        <TextField label="Confirmer le mot de passe" type="password" autoComplete="new-password" required value={confirmation} onChange={e => setConfirmation(e.target.value)} disabled={busy} />
      </>}
      {error && <Alert severity="error">{error}</Alert>}{message && <Alert severity="success">{message}</Alert>}
      <Button disabled={busy} onClick={() => { setMode(mode === 'request' ? 'reset' : 'request'); setError(''); setMessage('') }}>{mode === 'request' ? 'J’ai déjà un jeton' : 'Demander un lien'}</Button>
    </Stack></DialogContent><DialogActions><Button disabled={busy} onClick={onClose}>Fermer</Button><Button type="submit" variant="contained" disabled={busy || (local && mode === 'request')}>{mode === 'request' ? 'Demander un lien' : 'Réinitialiser'}</Button></DialogActions></form>
  </Dialog>
}
