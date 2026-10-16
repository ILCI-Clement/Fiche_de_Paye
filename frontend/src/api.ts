import type { Filter, LeaveDetail, LeaveInput, LeaveList, LeaveRequest, Scope, Session } from './types'

export class ApiError extends Error {
  status: number
  constructor(message: string, status: number) { super(message); this.status = status }
}

export async function request<T>(path: string, token: string | null, init?: RequestInit, options?: { blob?: boolean; timeout?: number }): Promise<T> {
  let response: Response
  try {
    response = await fetch(`/api${path}`, {
      ...init,
      signal: AbortSignal.any([AbortSignal.timeout(options?.timeout ?? 15000), ...(init?.signal ? [init.signal] : [])]),
      headers: { 'Content-Type': 'application/json', ...(token ? { Authorization: `Bearer ${token}` } : {}), ...init?.headers },
    })
  } catch (error) {
    if (error instanceof DOMException && error.name === 'AbortError') throw error
    throw new ApiError('Le serveur est indisponible. Réessayez dans un instant.', 0)
  }
  if (response.ok && options?.blob) {
    const disposition = response.headers.get('Content-Disposition') || ''
    const encoded = disposition.match(/filename\*=UTF-8''(.+)$/)?.[1]
    const filename = encoded ? decodeURIComponent(encoded) : disposition.match(/filename="([^"]+)"/)?.[1] || 'presence.zip'
    return { blob: await response.blob(), filename } as T
  }
  const body = await response.json().catch(() => null)
  if (!response.ok) {
    if (response.status === 401 && token) window.dispatchEvent(new Event('presence:expired'))
    const detail = body?.detail
    const message = typeof detail === 'string' ? detail : Array.isArray(detail)
      ? detail.map((item: { msg: string }) => item.msg.replace(/^Value error, /, '')).join(' ')
      : 'Une erreur est survenue. Veuillez réessayer.'
    throw new ApiError(message, response.status)
  }
  if (body === null) throw new ApiError('Réponse du serveur invalide.', response.status)
  return body as T
}

export const api = {
  session: (token: string, signal?: AbortSignal) => request<Session>('/leave/session', token, { signal }),
  login: (username: string, password: string, remember_me = false) => request<{ auth_token: string; remember_token?: string }>('/login', null, { method: 'POST', body: JSON.stringify({ username, password, remember_me }) }),
  list: (token: string, scope: Scope, status: Filter, page: number, signal?: AbortSignal) =>
    request<LeaveList>(`/leave-requests?scope=${scope}&status=${status}&page=${page}`, token, { signal }),
  create: (token: string, data: LeaveInput) => request<LeaveRequest>('/leave-requests', token, { method: 'POST', body: JSON.stringify(data) }),
  detail: (token: string, id: number, signal?: AbortSignal) => request<LeaveDetail>(`/leave-requests/${id}`, token, { signal }),
  withdraw: (token: string, id: number) => request<LeaveRequest>(`/leave-requests/${id}/withdraw`, token, { method: 'POST' }),
  decide: (token: string, id: number, decision: 'approved' | 'rejected', comment: string) =>
    request<LeaveRequest>(`/leave-requests/${id}/decision`, token, { method: 'POST', body: JSON.stringify({ decision, comment }) }),
}
