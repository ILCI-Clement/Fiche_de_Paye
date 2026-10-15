import { request } from '../api'
export interface Deadline {
  reminder_key: string; employee_username: string; employee_name: string; contract_end_date: string;
  days_remaining: number; manager_name?: string; manager_username?: string;
}
export interface Interview {
  employee_username: string; interview_year: number; due_date: string | null;
  notes: string | null; completed: boolean; completed_at: string | null; completed_by: string | null;
}
export interface Dashboard { year: number; contract_reminders_auto_send: boolean; contracts: Deadline[]; interviews: Interview[] }
export const dashboardApi = {
  get: (token: string, signal?: AbortSignal) => request<Dashboard>('/dashboard', token, { signal }),
  setting: (token: string, auto_send: boolean) => request('/settings/contract-reminders', token, { method: 'PATCH', body: JSON.stringify({ auto_send }) }),
  send: (token: string) => request('/contract-reminders/send', token, { method: 'POST' }),
  resend: (token: string, item: Deadline) => request('/contract-reminders/resend', token, { method: 'POST', body: JSON.stringify({ reminder_key: item.reminder_key, contract_end_date: item.contract_end_date }) }),
  interview: (token: string, item: Interview, data: { due_date: string | null; notes: string; completed: boolean }) => request(`/annual-interviews/${encodeURIComponent(item.employee_username)}/${item.interview_year}`, token, { method: 'PATCH', body: JSON.stringify(data) }),
}
