import { useQuery } from '@tanstack/react-query'
import { api } from './api'
import type { Session } from './types'

// The shell owns polling; the leave page shares the same query and invalidation.
export function useLeaveReviews(token: string, session: Session, poll = false) {
  return useQuery({
    queryKey: ['leave-list', session.username, 'review-count'],
    queryFn: ({ signal }) => api.list(token, 'team', 'pending', 1, signal),
    enabled: session.can_review,
    staleTime: 15000,
    refetchInterval: poll && session.can_review ? 30000 : false,
    retry: false,
  })
}
