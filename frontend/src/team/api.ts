import { request } from '../api'

export interface Group { id: number; name: string; is_active?: boolean }
export interface Department { id: number; name: string; leaders: string[]; members: string[] }
export interface Organization {
  general_id: number | null; departments: Department[]; people_count: number;
  unassigned: { Nom: string; Situation: string }[]; issues: string[];
}
export interface TeamMember { name: string; group_ids: number[]; managed_group_ids: number[] }
export interface DirectoryPerson { username: string; display_name: string; group_ids: number[]; managed_group_ids: number[] }
export interface Personnel {
  username: string; email: string; employee_type?: string; manager_id?: string | null;
  group_ids: number[]; role_tags?: string[]; role?: string; is_admin?: boolean;
}
export const teamApi = {
  directory: (token: string, signal?: AbortSignal) => request<{ organization: Organization; groups: Group[]; users: DirectoryPerson[] }>('/organization-directory', token, { signal }),
  mine: (token: string, signal?: AbortSignal) => request<{ groups: Group[]; members: TeamMember[] }>('/my-team', token, { signal }),
  people: (token: string, signal?: AbortSignal) => request<{ users: Personnel[] }>('/list-users', token, { signal }),
  groups: (token: string, signal?: AbortSignal) => request<{ groups: Group[] }>('/groups', token, { signal }),
}
