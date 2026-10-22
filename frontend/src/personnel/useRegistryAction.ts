import { useRef, useState } from 'react'
import { useQueryClient } from '@tanstack/react-query'
import { ApiError } from '../api'
import { useDraftGuard } from '../draftGuard'

/** A failed or stale write never resets a form or silently retries its contents. */
export function useRegistryAction() {
  const client = useQueryClient()
  const lock = useRef(false)
  const [busy, setBusy] = useState(false)
  const [error, setError] = useState('')
  const [conflict, setConflict] = useState(false)
  useDraftGuard(false, busy, 'Registre du personnel')
  async function run(action: () => Promise<unknown>) {
    if (lock.current) return false
    lock.current = true; setBusy(true); setError(''); setConflict(false)
    try {
      await action()
      await Promise.all([
        client.invalidateQueries({ queryKey: ['personnel-register'] }),
        client.invalidateQueries({ queryKey: ['account-personnel'] }),
        client.invalidateQueries({ queryKey: ['account-personnel-choices'] }),
      ])
      return true
    } catch (value) {
      const stale = value instanceof ApiError && value.status === 409
      setConflict(stale)
      setError(value instanceof Error ? value.message : 'Impossible d’enregistrer les modifications.')
      return false
    } finally { lock.current = false; setBusy(false) }
  }
  return { busy, error, conflict, run }
}
