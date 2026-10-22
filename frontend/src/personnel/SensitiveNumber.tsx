import { useCallback, useEffect, useLayoutEffect, useRef, useState } from 'react'
import { Alert, Box, Button, CircularProgress, Typography } from '@mui/material'
import VisibilityOutlined from '@mui/icons-material/VisibilityOutlined'
import VisibilityOffOutlined from '@mui/icons-material/VisibilityOffOutlined'
import { useQueryClient } from '@tanstack/react-query'
import { ApiError } from '../api'
import { useWorkspaceToolbar } from '../workspaceToolbar'
import { personnelApi } from './api'
import { useRegistryViewer } from './context'
import { useSensitiveNumberCapability } from './useSensitiveNumberCapability'

/** Plaintext lives only in this short-lived view, never in query cache or storage. */
export function SensitiveNumber({ token, employeeId, hasNumber, active = true }: { token: string; employeeId: number; hasNumber: boolean; active?: boolean }) {
  const capability = useSensitiveNumberCapability(token)
  const viewer = useRegistryViewer()
  const slot = useWorkspaceToolbar()
  // Mount a fresh short-lived state for every identity / capability. Hiding the page
  // unmounts it synchronously instead of retaining a number in an invisible effect.
  if (!active || slot?.active === false) return null
  const enabled = capability.data?.nir_enabled === true && !capability.error
  return <SensitiveNumberView key={JSON.stringify([viewer, employeeId, token, hasNumber, enabled])}
    token={token} employeeId={employeeId} hasNumber={hasNumber} viewer={viewer}
    enabled={enabled} pending={capability.isPending} checking={capability.isFetching} />
}

function SensitiveNumberView({ token, employeeId, hasNumber, viewer, enabled, pending, checking }: {
  token: string; employeeId: number; hasNumber: boolean; viewer: string; enabled: boolean; pending: boolean; checking: boolean;
}) {
  const client = useQueryClient()
  const [value, setValue] = useState('')
  const [busy, setBusy] = useState(false)
  const [error, setError] = useState('')
  const request = useRef<AbortController | null>(null)
  const timer = useRef<ReturnType<typeof setTimeout> | undefined>(undefined)
  const generation = useRef(0)
  const invalidate = useCallback(() => {
    generation.current += 1
    request.current?.abort(); request.current = null
    if (timer.current !== undefined) { clearTimeout(timer.current); timer.current = undefined }
  }, [])
  const hide = useCallback(() => {
    invalidate(); setValue(''); setError(''); setBusy(false)
  }, [invalidate])
  // Unmount/key changes synchronously abort pending access before a later response
  // could be applied; the plaintext-bearing component itself is discarded.
  useLayoutEffect(() => invalidate, [invalidate])
  useEffect(() => {
    const onVisibility = () => { if (document.visibilityState === 'hidden') hide() }
    document.addEventListener('visibilitychange', onVisibility)
    return () => document.removeEventListener('visibilitychange', onVisibility)
  }, [hide])
  async function reveal() {
    if (busy || request.current || !hasNumber || !enabled || document.visibilityState !== 'visible') return
    invalidate()
    const current = generation.current
    const controller = new AbortController()
    request.current = controller
    setError(''); setBusy(true)
    try {
      // Deliberately not useQuery/useMutation: the response must never enter their cache.
      const result = await personnelApi.revealSocialSecurityNumber(token, employeeId, controller.signal)
      if (controller.signal.aborted || current !== generation.current || document.visibilityState !== 'visible') return
      if (typeof result.social_security_number !== 'string' || !result.social_security_number) throw new Error('Unavailable')
      setValue(result.social_security_number)
      timer.current = setTimeout(hide, 30000)
      // Only the metadata-only audit query is invalidated; plaintext is never cached.
      void client.invalidateQueries({ queryKey: ['personnel-register', viewer, 'history', employeeId] })
    } catch (reason) {
      if (controller.signal.aborted || current !== generation.current) return
      // Never echo a server response or validation value into this sensitive display.
      setError(reason instanceof ApiError && reason.status === 403 ? 'Vous n’avez plus l’autorisation de consulter ce numéro.' : 'Impossible d’afficher le numéro. Actualisez la fiche puis réessayez.')
    } finally {
      if (current === generation.current) { request.current = null; setBusy(false) }
    }
  }
  return <Box>
    <Typography variant="body2" color="text.secondary">{hasNumber ? 'Renseigné · stockage chiffré' : 'Non renseigné'}</Typography>
    {value && <Typography component="output" aria-label="Numéro de sécurité sociale affiché temporairement" sx={{ display: 'block', mt: 1, fontFamily: 'monospace', overflowWrap: 'anywhere', userSelect: 'text' }}>{value}</Typography>}
    {error && <Alert severity="error" sx={{ mt: 1 }}>{error}</Alert>}
    {hasNumber && <Box sx={{ display: 'flex', alignItems: 'center', flexWrap: 'wrap', gap: 1, mt: 1 }}>
      {value ? <Button size="small" startIcon={<VisibilityOffOutlined />} onClick={hide}>Masquer</Button> : busy ? <><CircularProgress size={18} aria-label="Chargement du numéro protégé" /><Button size="small" onClick={hide}>Annuler</Button></> : <Button size="small" startIcon={<VisibilityOutlined />} disabled={!enabled || checking} onClick={() => void reveal()}>Afficher pendant 30 secondes</Button>}
    </Box>}
    {hasNumber && <Typography variant="caption" color="text.secondary">La consultation est enregistrée. Le numéro est masqué après 30 secondes ou en quittant cette vue.</Typography>}
    {hasNumber && !pending && !enabled && <Typography variant="caption" color="text.secondary" sx={{ display: 'block', mt: .5 }}>La consultation protégée est indisponible dans cet environnement.</Typography>}
  </Box>
}
