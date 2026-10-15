import { useDraftGuard, useWorkspaceGuard } from '../draftGuard'
import { useRef, useState } from 'react'
import { Alert, Box, Button, CircularProgress, Paper, TextField, Typography } from '@mui/material'
import { useQuery } from '@tanstack/react-query'
import { request } from '../api'
import type { Session } from '../types'

export function ProfilePage({ token, session, onSession }: { token: string; session: Session; onSession: (token: string) => void }) {
  const profile = useQuery({ queryKey: ['own-profile', session.username], queryFn: ({ signal }) => request<{ users: { username: string; email: string }[] }>('/list-users', token, { signal }), retry: false })
  const current = profile.data?.users.find(p => p.username === session.username)
  return <Box className="workspace"><Typography variant="h4" component="h1" sx={{ mb: 3 }}>Paramètres du compte</Typography>
    {profile.isPending ? <CircularProgress /> : profile.error ? <Alert severity="error">{profile.error.message}<Button onClick={() => void profile.refetch()}>Réessayer</Button></Alert> : current ?
      <ProfileForm key={current.username} token={token} username={current.username} email={current.email} onSession={onSession} /> : <Alert severity="error">Votre compte est introuvable.</Alert>}
  </Box>
}

function ProfileForm({ token, username, email, onSession }: { token: string; username: string; email: string; onSession: (token: string) => void }) {
  const guard = useWorkspaceGuard()
  const [name, setName] = useState(username)
  const [mail, setMail] = useState(email)
  const [password, setPassword] = useState('')
  const [confirm, setConfirm] = useState('')
  const [current, setCurrent] = useState('')
  const [busy, setBusy] = useState(false)
  const lock = useRef(false)
  const [error, setError] = useState('')
  useDraftGuard(name !== username || mail !== email || !!password || !!confirm || !!current, busy, 'Paramètres du compte')
  async function save() {
    if (lock.current) return
    if (password !== confirm) { setError('Les nouveaux mots de passe ne correspondent pas.'); return }
    lock.current = true; setBusy(true); setError('')
    try {
      const result = await request<{ auth_token: string }>(`/update_profile/${encodeURIComponent(username)}`, token, { method: 'PUT', body: JSON.stringify({ current_password: current, new_username: name.trim(), new_email: mail.trim(), new_password: password || null, confirm_password: confirm || null }) })
      setPassword(''); setConfirm(''); setCurrent(''); onSession(result.auth_token)
    } catch (e) { setError(e instanceof Error ? e.message : 'Impossible de modifier le compte.') }
    finally { lock.current = false; setBusy(false) }
  }
  return <Paper component="form" sx={{ p: { xs: 2, md: 3 }, maxWidth: 760 }} onSubmit={e => { e.preventDefault(); if (guard) guard.request(() => void save(), undefined, 'Paramètres du compte'); else void save() }}>
    <Alert severity="info" sx={{ mb: 3 }}>Enregistrez d’abord vos autres modifications. La mise à jour du compte renouvelle votre session et recharge votre espace.</Alert>
    <Box sx={{ display: 'grid', gridTemplateColumns: { xs: '1fr', md: '1fr 1fr' }, gap: 2 }}>
      <TextField label="Nom d’utilisateur" autoComplete="username" required value={name} disabled={busy} onChange={e => setName(e.target.value)} />
      <TextField label="Adresse e-mail" type="email" autoComplete="email" required value={mail} disabled={busy} onChange={e => setMail(e.target.value)} />
      <TextField label="Nouveau mot de passe (optionnel)" type="password" autoComplete="new-password" value={password} disabled={busy} onChange={e => setPassword(e.target.value)} slotProps={{ htmlInput: { minLength: 8 } }} />
      <TextField label="Confirmer le nouveau mot de passe" type="password" autoComplete="new-password" value={confirm} disabled={busy} onChange={e => setConfirm(e.target.value)} />
      <TextField label="Mot de passe actuel" type="password" autoComplete="current-password" required value={current} disabled={busy} onChange={e => setCurrent(e.target.value)} />
      <Button type="submit" variant="contained" disabled={busy || !name.trim() || !mail.trim()}>Enregistrer les modifications</Button>
    </Box>{error && <Alert severity="error" sx={{ mt: 2 }}>{error}</Alert>}
  </Paper>
}
