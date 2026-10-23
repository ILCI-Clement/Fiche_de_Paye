import { useCapabilities } from '../capabilities'
import { useEffect, useState } from 'react'
import { Alert, Box, Button, Checkbox, FormControlLabel, Paper, Stack, TextField, Typography } from '@mui/material'
import { useMutation } from '@tanstack/react-query'
import { api, ApiError } from '../api'
import { ThemeToggle } from './ThemeToggle'
import { PasswordRecovery } from './PasswordRecovery'
import { storeRememberedToken } from '../authSession'
import { brand } from '../brand'

export function Login({ onLogin }: { onLogin: (token: string) => void }) {
  const [username, setUsername] = useState('')
  const [password, setPassword] = useState('')
  const capabilities = useCapabilities()
  const [remember, setRemember] = useState(false)
  const [resetToken, setResetToken] = useState(() => new URLSearchParams(window.location.search).get('token') || '')
  const [recovery, setRecovery] = useState(Boolean(resetToken))
  const [cooldown, setCooldown] = useState<{ identifier: string; until: number } | null>(null)
  const [clock, setClock] = useState(Date.now)
  useEffect(() => {
    if (!cooldown) return
    const interval = window.setInterval(() => {
      const now = Date.now()
      setClock(now)
      if (now >= cooldown.until) window.clearInterval(interval)
    }, 1000)
    return () => window.clearInterval(interval)
  }, [cooldown])
  const seconds = cooldown?.identifier === username.trim().toLowerCase() ? Math.max(0, Math.ceil((cooldown.until - clock) / 1000)) : 0
  const remaining = `${Math.floor(seconds / 60).toString().padStart(2, '0')}:${(seconds % 60).toString().padStart(2, '0')}`
  useEffect(() => {
    const url = new URL(window.location.href)
    if (url.searchParams.has('token')) { url.searchParams.delete('token'); window.history.replaceState(null, '', url.pathname + url.search + url.hash) }
  }, [])
  const login = useMutation({ mutationFn: () => api.login(username, password, remember),
    onError: error => {
      if (error instanceof ApiError && error.status === 429 && error.retryAfterSeconds) {
        const now = Date.now()
        setClock(now); setCooldown({ identifier: username.trim().toLowerCase(), until: now + error.retryAfterSeconds * 1000 })
      }
    },
    onSuccess: data => { storeRememberedToken(data.remember_token); setPassword(''); setCooldown(null); onLogin(data.auth_token) } })
  return <Box className="login-layout">
    <Box className="login-intro">
      <Box className="brand-mark">{brand.mark}</Box>
      <Typography variant="overline" sx={{ color: 'var(--sidebar-muted)', letterSpacing: 3, mt: 5 }}>{brand.title}</Typography>
      <Typography component="h1" sx={{ fontSize: { xs: 32, md: 48 }, fontWeight: 750, lineHeight: 1.15, my: 2 }}>Vos démarches RH.<br />Un espace unique.</Typography>
      <Typography sx={{ color: 'var(--sidebar-muted)', maxWidth: 390, lineHeight: 1.8 }}>Retrouvez vos informations RH, vos contrats et documents, gérez vos présences et suivez vos demandes de congé.</Typography>
      <Box sx={{ mt: 'auto', pt: 5, color: 'var(--sidebar-muted)', fontSize: 13 }}>Personnel · Contrats · Documents · Congés</Box>
    </Box>
    <Box className="login-content"><Paper sx={{ width: '100%', maxWidth: 470, p: { xs: 3, md: 4 }, border: '1px solid', borderColor: 'divider' }}>
      <Box sx={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}><Typography variant="h4">Bienvenue</Typography><ThemeToggle /></Box>
      <Typography color="text.secondary" sx={{ mt: 1, mb: 3 }}>Connectez-vous à votre {brand.name}.</Typography>
      {capabilities.data?.environment === 'local-snapshot' && <Alert severity="info" sx={{ mb: 3 }}>Version locale de test. Utilisez votre compte habituel ; les modifications restent dans la copie locale.</Alert>}
      {seconds > 0 ? <Alert severity="warning" sx={{ mb: 2 }}>Connexion temporairement suspendue après 5 mots de passe incorrects. Réessayez dans {remaining}.</Alert> : login.error && !(login.error instanceof ApiError && login.error.status === 429 && login.error.retryAfterSeconds) && <Alert severity="error" sx={{ mb: 2 }}>{login.error.message}</Alert>}
      <Box component="form" onSubmit={event => { event.preventDefault(); if (!seconds && !login.isPending) login.mutate() }}>
        <Stack spacing={2.5}>
          <TextField label="Adresse e-mail ou nom de compte" autoComplete="username" autoCapitalize="none" spellCheck={false} value={username} disabled={login.isPending} onChange={event => setUsername(event.target.value)} required fullWidth />
          <TextField label="Mot de passe" type="password" autoComplete="current-password" value={password} disabled={login.isPending} onChange={event => setPassword(event.target.value)} required fullWidth />
          <Button type="submit" variant="contained" disabled={login.isPending || seconds > 0 || !username || !password}>{seconds > 0 ? `Réessayer dans ${remaining}` : 'Se connecter'}</Button>
          <FormControlLabel control={<Checkbox checked={remember} disabled={login.isPending} onChange={e => setRemember(e.target.checked)} />} label="Se souvenir de moi sur cet appareil (30 jours)" />
          <Button disabled={login.isPending} onClick={() => setRecovery(true)}>Mot de passe oublié ?</Button>
        </Stack>
      </Box>
    </Paper></Box>
    {recovery && <PasswordRecovery initialToken={resetToken} onClose={() => { setRecovery(false); setResetToken('') }} />}
  </Box>
}
