import { useEffect, useState } from 'react'
import { Alert, Box, Button, CircularProgress, Snackbar, Typography } from '@mui/material'
import { useQuery, useQueryClient } from '@tanstack/react-query'
import { api, ApiError } from './api'
import { forgetRememberedToken, rememberedToken, restoreRememberedToken, storeRememberedToken } from './authSession'
import { Login } from './components/Login'
import { AppWorkspace } from './AppWorkspace'

const SESSION_KEY = 'presence-react-session'

export default function App() {
  const client = useQueryClient()
  const [token, setToken] = useState<string | null>(() => sessionStorage.getItem(SESSION_KEY))
  const [expired, setExpired] = useState(false)
  const [restoring, setRestoring] = useState(() => !sessionStorage.getItem(SESSION_KEY) && Boolean(rememberedToken()))
  const [restoreError, setRestoreError] = useState('')
  useEffect(() => {
    if (sessionStorage.getItem(SESSION_KEY)) return
    const saved = rememberedToken()
    if (!saved) return
    let disposed = false
    restoreRememberedToken(saved).then(result => {
      if (!disposed) { sessionStorage.setItem(SESSION_KEY, result.auth_token); setToken(result.auth_token) }
    }).catch(error => {
      if (disposed) return
      if (error instanceof ApiError && error.status === 401) storeRememberedToken()
      else setRestoreError('Impossible de restaurer la session. Vous pouvez vous reconnecter.')
    }).finally(() => { if (!disposed) setRestoring(false) })
    return () => { disposed = true }
  }, [])
  function updateSession(value: string) {
    void forgetRememberedToken()
    sessionStorage.setItem(SESSION_KEY, value)
    client.clear()
    setToken(value)
    setExpired(false)
  }
  function logout() {
    void forgetRememberedToken()
    sessionStorage.removeItem(SESSION_KEY)
    setToken(null)
    client.clear()
  }
  useEffect(() => {
    let refreshing = false
    let disposed = false
    const expire = () => {
      if (refreshing) return
      refreshing = true
      const saved = rememberedToken()
      sessionStorage.removeItem(SESSION_KEY); setToken(null); client.clear()
      if (!saved) { setExpired(true); refreshing = false; return }
      setRestoring(true)
      restoreRememberedToken(saved).then(result => {
        if (!disposed) { sessionStorage.setItem(SESSION_KEY, result.auth_token); setToken(result.auth_token) }
      }).catch(error => {
        if (disposed) return
        if (error instanceof ApiError && error.status === 401) storeRememberedToken()
        setExpired(true)
      }).finally(() => { if (!disposed) { setRestoring(false); refreshing = false } })
    }
    window.addEventListener('presence:expired', expire)
    return () => { disposed = true; window.removeEventListener('presence:expired', expire) }
  }, [client])
  const session = useQuery({ queryKey: ['session', token], queryFn: ({ signal }) => api.session(token!, signal), enabled: Boolean(token), retry: false, refetchInterval: 30000 })
  if (restoring) return <Box className="center-state"><CircularProgress /><Typography>Restauration de votre session…</Typography></Box>
  if (!token) return <>{restoreError && <Alert severity="info" onClose={() => setRestoreError('')}>{restoreError}</Alert>}<Login onLogin={value => { sessionStorage.setItem(SESSION_KEY, value); setToken(value); setExpired(false); setRestoreError('') }} />
    <Snackbar open={expired} autoHideDuration={5000} onClose={() => setExpired(false)} message="Votre session a expiré. Reconnectez-vous." /></>
  if (session.isPending) return <Box className="center-state"><CircularProgress /><Typography>Ouverture de votre espace…</Typography></Box>
  if (session.error) return <Box className="center-state"><Alert severity="error">{session.error.message}</Alert>
    <Button onClick={() => void session.refetch()}>Réessayer</Button><Button onClick={logout}>Retour à la connexion</Button></Box>
  if (!session.data) return null
  return <AppWorkspace key={session.data.username} token={token} session={session.data} logout={logout} onSession={updateSession} />
}

