import { request } from '../api'

export type PersonnelType = '' | 'salarie' | 'stagiaire' | 'alternance'
export type ContractType = '' | 'CDI' | 'CDD' | 'stage' | 'autre'
export type EmploymentStatus = 'employed' | 'left' | 'upcoming' | 'unknown'
export interface EmploymentInput {
  entry_date: string | null; exit_date: string | null; contract_type: ContractType; other_contract: string;
}
export interface EmploymentPeriod extends EmploymentInput { period_id: number }
export interface PersonnelInput {
  full_name: string; nationality: string; sex: '' | 'F' | 'M'; address: string;
  birth_date: string | null; job_title: string; qualification: string;
  personnel_type: PersonnelType; linked_username: string | null;
}
/** Write-only value: never part of an employee DTO, cached detail or audit snapshot. */
export type PersonnelWriteInput = PersonnelInput & { social_security_number?: string | null }
export interface Employee extends PersonnelInput {
  linked_display_name?: string | null;
  employee_id: number; registration_number: number; is_archived: boolean; version: number;
  has_social_security_number: boolean;
  latest_period: EmploymentPeriod | null; status: EmploymentStatus; created_at: string; updated_at: string;
}
export interface PersonnelDetail { employee: Employee; periods: EmploymentPeriod[] }
export interface PersonnelList {
  employees: Employee[]; total: number; page: number; page_size: number;
}
export interface PersonnelFilters {
  q: string; personnel_type: PersonnelType; archive: 'active' | 'archived' | 'all';
  status: 'all' | EmploymentStatus; page: number; page_size: number;
}
export interface RegistryAccount {
  username: string; email: string; is_active: boolean; linked_employee_id: number | null;
  full_name?: string | null; display_name?: string;
}
export interface RegistrySnapshot { employee: Partial<Employee>; periods: EmploymentPeriod[] }
export interface RegistryEvent {
  event_id: number; actor_username: string; action: string; created_at: string;
  actor_display_name?: string;
  before: RegistrySnapshot | null; after: RegistrySnapshot;
}
const path = '/personnel-register'
const body = (method: string, value: unknown) => ({ method, body: JSON.stringify(value) })
export const personnelApi = {
  capabilities: (token: string, signal?: AbortSignal) => request<{ nir_enabled: boolean }>(`${path}/capabilities`, token, { signal, cache: 'no-store' }),
  list: (token: string, filters: PersonnelFilters, signal?: AbortSignal) => {
    const params = new URLSearchParams(Object.entries(filters).map(([key, value]) => [key, String(value)]))
    return request<PersonnelList>(`${path}?${params}`, token, { signal })
  },
  detail: (token: string, id: number, signal?: AbortSignal) => request<PersonnelDetail>(`${path}/${id}`, token, { signal }),
  create: (token: string, data: PersonnelWriteInput & { initial_period?: EmploymentInput }) => request<PersonnelDetail>(path, token, body('POST', data)),
  update: (token: string, id: number, data: PersonnelWriteInput, version: number) => request<PersonnelDetail>(`${path}/${id}`, token, body('PUT', { ...data, expected_version: version })),
  revealSocialSecurityNumber: (token: string, id: number, signal: AbortSignal) => request<{ social_security_number: string }>(`${path}/${id}/social-security-number/reveal`, token, { method: 'POST', signal, cache: 'no-store' }),
  archive: (token: string, id: number, archived: boolean, version: number) => request<PersonnelDetail>(`${path}/${id}/archive`, token, body('PATCH', { is_archived: archived, expected_version: version })),
  remove: (token: string, id: number, version: number, confirmation: string) => request(`${path}/${id}`, token, body('DELETE', { expected_version: version, confirmation })),
  savePeriod: (token: string, id: number, periodId: number | null, data: EmploymentInput, version: number) => request<PersonnelDetail>(`${path}/${id}/periods${periodId === null ? '' : `/${periodId}`}`, token, body(periodId === null ? 'POST' : 'PUT', { ...data, expected_version: version })),
  history: (token: string, id: number, page: number, signal?: AbortSignal) => request<{ events: RegistryEvent[]; total: number; page: number; page_size: number }>(`${path}/${id}/history?page=${page}&page_size=20`, token, { signal }),
  accounts: (token: string, search: string, signal?: AbortSignal) => request<{ accounts: RegistryAccount[]; has_more: boolean; limit: number }>(`${path}/accounts?${new URLSearchParams({ q: search })}`, token, { signal }),
}

export const personnelTypes: { value: PersonnelType; label: string }[] = [
  { value: '', label: 'Non renseigné' }, { value: 'salarie', label: 'Salarié' },
  { value: 'stagiaire', label: 'Stagiaire' }, { value: 'alternance', label: 'Alternant' },
]
export const contractTypes: { value: ContractType; label: string }[] = [
  { value: '', label: 'Non renseigné' }, { value: 'CDI', label: 'CDI' }, { value: 'CDD', label: 'CDD' },
  { value: 'stage', label: 'Convention de stage' }, { value: 'autre', label: 'Autre' },
]
export const statusLabels: Record<EmploymentStatus, string> = {
  employed: 'En poste', left: 'Sorti', upcoming: 'À venir', unknown: 'À préciser',
}
export const typeLabel = (value: PersonnelType) => personnelTypes.find(item => item.value === value)?.label || 'Non renseigné'
export const contractLabel = (period?: EmploymentInput | null) => !period ? '—' : period.contract_type === 'autre' ? period.other_contract || 'Autre' : contractTypes.find(item => item.value === period.contract_type)?.label || 'Non renseigné'
export const dateLabel = (value?: string | null) => !value ? '—' : /^\d{4}-\d{2}-\d{2}$/.test(value) ? value.split('-').reverse().join('/') : value
export const blankPeriod = (): EmploymentInput => ({ entry_date: null, exit_date: null, contract_type: '', other_contract: '' })
