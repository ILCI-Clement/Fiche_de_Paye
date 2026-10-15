import { request } from './api'
const REMEMBER_KEY = 'presence-react-remember'
export interface LoginResult { auth_token: string; remember_token?: string }
export function rememberedToken() {
  try { return localStorage.getItem(REMEMBER_KEY) || '' } catch { return '' }
}
export function storeRememberedToken(token?: string) {
  try { if (token) localStorage.setItem(REMEMBER_KEY, token); else localStorage.removeItem(REMEMBER_KEY) } catch { /* Session-only login remains available. */ }
}
export async function forgetRememberedToken() {
  const token = rememberedToken()
  storeRememberedToken()
  if (token) await request('/forget-session', null, { method: 'POST', body: JSON.stringify({ remember_token: token }) }).catch(() => undefined)
}
export function restoreRememberedToken(token: string) {
  return request<LoginResult>('/restore-session', null, { method: 'POST', body: JSON.stringify({ remember_token: token }) })
}
