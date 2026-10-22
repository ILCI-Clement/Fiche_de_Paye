import { request } from '../api'
import type { Group } from '../team/api'

export interface AdminUser {
  is_active?: boolean; login_deleted_at?: string | null;
  username: string; email: string; role_tags: string[]; employee_type: string;
  full_name?: string | null; display_name?: string; manager_display_name?: string; job_title?: string | null;
  manager_id: string | null; group_ids: number[]; managed_group_ids: number[];
}
export interface OrganizationInput {
  role_tags: string[]; employee_type: string; manager_id: string | null;
  group_ids: number[]; managed_group_ids: number[];
}
export interface ContractInput { employee_type: string; start_date: string; end_date: string | null; is_cdi: boolean }
export interface Contract extends ContractInput { id: number; is_archived?: boolean }
const userPath = (name: string) => `/users/${encodeURIComponent(name)}`
export const adminApi = {
  users: (token: string, signal?: AbortSignal) => request<{ users: AdminUser[] }>('/list-users', token, { signal }),
  groups: (token: string, signal?: AbortSignal) => request<{ groups: Group[] }>('/groups', token, { signal }),
  organization: (token: string, name: string, data: OrganizationInput) => request(`${userPath(name)}/organization`, token, { method: 'PUT', body: JSON.stringify(data) }),
  manager: (token: string, name: string, manager: string) => request(`${userPath(name)}/direct-manager`, token, { method: 'PATCH', body: JSON.stringify({ manager_id: manager }) }),
  activation: (token: string, name: string, is_active: boolean) => request(`${userPath(name)}/activation`, token, { method: 'PATCH', body: JSON.stringify({ is_active }) }),
  deleteLogin: (token: string, name: string, confirmation: string) => request(`${userPath(name)}/login`, token, { method: 'DELETE', body: JSON.stringify({ confirmation }) }),
  purge: (token: string, name: string, confirmation: string, reason: string) => request(`/delete-user/${encodeURIComponent(name)}`, token, { method: 'DELETE', body: JSON.stringify({ confirmation, reason }) }),
  activationHistory: (token: string, name: string, signal?: AbortSignal) => request<{ events: { actor_username: string; actor_display_name?: string; was_active: boolean; is_active: boolean; action?: string; created_at: string }[] }>(`${userPath(name)}/activation-history`, token, { signal }),
  saveGroup: (token: string, id: number | null, name: string, is_active: boolean) => request(id === null ? '/groups' : `/groups/${id}`, token, { method: id === null ? 'POST' : 'PUT', body: JSON.stringify({ name, is_active }) }),
  contracts: (token: string, name: string, signal?: AbortSignal) => request<{ contracts: Contract[] }>(`${userPath(name)}/contracts`, token, { signal }),
  saveContract: (token: string, name: string, id: number | null, data: ContractInput) => request(`${userPath(name)}/contracts${id === null ? '' : `/${id}`}`, token, { method: id === null ? 'POST' : 'PATCH', body: JSON.stringify(data) }),
  archiveContract: (token: string, name: string, id: number, is_archived: boolean) => request(`${userPath(name)}/contracts/${id}/archive`, token, { method: 'PATCH', body: JSON.stringify({ is_archived }) }),
  deleteContract: (token: string, name: string, id: number, confirmation: string) => request(`${userPath(name)}/contracts/${id}`, token, { method: 'DELETE', body: JSON.stringify({ confirmation }) }),
}
export type RunAction = (action: () => Promise<unknown>, message: string) => Promise<boolean>
export const personnelTypes = [['salarie', 'Salarié'], ['stagiaire', 'Stagiaire'], ['alternance', 'Alternance']]
