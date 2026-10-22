import { useQuery } from '@tanstack/react-query'
import { personnelApi } from './api'
import { useRegistryViewer } from './context'

/** Only a capability flag is cached; never a social-security number. */
export function useSensitiveNumberCapability(token: string) {
  const viewer = useRegistryViewer()
  return useQuery({ queryKey: ['personnel-register', viewer, 'capabilities'], queryFn: ({ signal }) => personnelApi.capabilities(token, signal), retry: false })
}
