import { useQuery } from '@tanstack/react-query'
import { request } from './api'

interface Capabilities { environment: 'local-snapshot' | 'server'; account_creation: boolean; mail_enabled: boolean }
export function useCapabilities() {
  const query = useQuery({ queryKey: ['ui-capabilities'], queryFn: ({ signal }) => request<Capabilities>('/ui-capabilities', null, { signal }), retry: false, staleTime: 60000 })
  // A missing or unreachable capability endpoint never enables an external action.
  return { ...query, canCreateAccounts: query.data?.account_creation === true, canSendMail: query.data?.mail_enabled === true }
}
