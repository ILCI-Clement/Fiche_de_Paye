export type AccountStatus = 'active' | 'inactive' | 'deleted' | 'archived'
export type AccountStatusFilter = AccountStatus | 'all'

export function accountStatus(account: { is_active?: boolean; login_deleted_at?: string | null; personnel_archived?: boolean }): AccountStatus {
  return account.login_deleted_at ? 'deleted' : account.is_active === false ? 'inactive' : account.personnel_archived ? 'archived' : 'active'
}
