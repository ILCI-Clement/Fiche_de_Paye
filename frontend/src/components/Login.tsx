import { useCapabilities } from '../capabilities'
import { useEffect, useState } from 'react'
import { Alert, Box, Button, Checkbox, FormControlLabel, Paper, Stack, TextField, Typography } from '@mui/material'
import { useMutation } from '@tanstack/react-query'
import { api } from '../api'
import { ThemeToggle } from './ThemeToggle'
import { PasswordRecovery } from './PasswordRecovery'
import { storeRememberedToken } from '../authSession'

export function Login({ onLogin }: { onLogin: (token: string) => void }) {
  const [username, setUsername] = useState('')
  const [password, setPassword] = useState('')
  const capabilities = useCapabilities()
  const [remember, setRemember] = useState(false)
  const [resetToken, setResetToken] = useState(() => new URLSearchParams(window.location.search).get('token') || '')
  const [recovery, setRecovery] = useState(Boolean(resetToken))
  useEffect(() => {
    const url = new URL(window.location.href)
    if (url.searchParams.has('token')) { url.searchParams.delete('token'); window.history.replaceState(null, '', url.pathname + url.search + url.hash) }
  }, [])
  const login = useMutation({ mutationFn: () => api.login(username, password, remember),
    onSuccess: data => { storeRememberedToken(data.remember_token); setPassword(''); onLogin(data.auth_token) } })
  return <Box className="login-layout">
    <Box className="login-intro">
      <Box className="brand-mark">P</Box>
      <Typography variant="overline" sx={{ color: 'var(--sidebar-muted)', letterSpacing: 3, mt: 5 }}>PRESENCE · ILCI</Typography>
      <Typography component="h1" sx={{ fontSize: { xs: 32, md: 48 }, fontWeight: 750, lineHeight: 1.15, my: 2 }}>Vos présences.<br />Votre espace de travail.</Typography>
      <Typography sx={{ color: 'var(--sidebar-muted)', maxWidth: 390, lineHeight: 1.8 }}>Retrouvez vos fiches de présence, générez vos documents et suivez vos demandes de congé.</Typography>
      <Box sx={{ mt: 'auto', pt: 5, color: 'var(--sidebar-muted)', fontSize: 13 }}>Présences · Congés payés · Espace RH</Box>
    </Box>
    <Box className="login-content"><Paper sx={{ width: '100%', maxWidth: 470, p: { xs: 3, md: 4 }, border: '1px solid', borderColor: 'divider' }}>
      <Box sx={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}><Typography variant="h4">Bienvenue</Typography><ThemeToggle /></Box>
      <Typography color="text.secondary" sx={{ mt: 1, mb: 3 }}>Connectez-vous avec votre compte Presence.</Typography>
      {capabilities.data?.environment === 'local-snapshot' && <Alert severity="info" sx={{ mb: 3 }}>Version locale de test. Utilisez votre compte habituel ; les modifications restent dans la copie locale.</Alert>}
      {login.error && <Alert severity="error" sx={{ mb: 2 }}>{login.error.message}</Alert>}
      <Box component="form" onSubmit={event => { event.preventDefault(); login.mutate() }}>
        <Stack spacing={2.5}>
          <TextField label="Nom de compte" autoComplete="username" value={username} onChange={event => setUsername(event.target.value)} required fullWidth />
          <TextField label="Mot de passe" type="password" autoComplete="current-password" value={password} onChange={event => setPassword(event.target.value)} required fullWidth />
          <Button type="submit" variant="contained" disabled={login.isPending || !username || !password}>Se connecter</Button>
          <FormControlLabel control={<Checkbox checked={remember} disabled={login.isPending} onChange={e => setRemember(e.target.checked)} />} label="Se souvenir de moi sur cet appareil (30 jours)" />
          <Button disabled={login.isPending} onClick={() => setRecovery(true)}>Mot de passe oublié ?</Button>
        </Stack>
      </Box>
    </Paper></Box>
    {recovery && <PasswordRecovery initialToken={resetToken} onClose={() => { setRecovery(false); setResetToken('') }} />}
  </Box>
}
