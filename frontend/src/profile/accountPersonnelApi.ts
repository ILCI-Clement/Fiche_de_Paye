import { request } from '../api'
import type { EmploymentPeriod, PersonnelInput } from '../personnel/api'

export interface LinkedPersonnel extends PersonnelInput {
  employee_id: number; registration_number: number; version: number; is_archived: boolean;
}
export interface AccountPersonnel {
  employee: LinkedPersonnel | null; periods: EmploymentPeriod[]; can_manage_link: boolean;
  account_display_name: string; account_email: string;
}
export interface PersonnelChoice {
  employee_id: number; full_name: string; job_title: string; personnel_type: string; version: number;
}
export interface PersonnelAssociation {
  employee_id: number | null; expected_employee_id: number | null;
  expected_version: number | null; target_version: number | null;
}
const path = (username: string) => `/account-personnel/${encodeURIComponent(username)}`
export const accountPersonnelApi = {
  get: (token: string, username: string, signal?: AbortSignal) => request<AccountPersonnel>(path(username), token, { signal, cache: 'no-store' }),
  choices: (token: string, username: string, search: string, signal?: AbortSignal) => request<{ employees: PersonnelChoice[]; has_more: boolean }>(`${path(username)}/choices?${new URLSearchParams({ q: search })}`, token, { signal, cache: 'no-store' }),
  link: (token: string, username: string, payload: PersonnelAssociation) => request<AccountPersonnel>(`${path(username)}/link`, token, { method: 'PUT', body: JSON.stringify(payload), cache: 'no-store' }),
}
