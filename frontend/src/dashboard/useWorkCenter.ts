import { useQuery } from '@tanstack/react-query'
import { useWorkspaceToolbar } from '../workspaceToolbar'
import { workCenterApi } from './workCenterApi'

export function useWorkCenter(token: string, viewer: string, month: number, year: number, allowed = true) {
  const slot = useWorkspaceToolbar()
  const active = allowed && slot?.active !== false
  return useQuery({ queryKey: ['work-center', viewer, year, month],
    queryFn: ({ signal }) => workCenterApi.get(token, month, year, signal),
    enabled: active, refetchInterval: active ? 60000 : false, staleTime: 15000, retry: false })
}
