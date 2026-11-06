export type Slot = 'full' | 'morning' | 'afternoon'
export type Status = 'pending' | 'approved' | 'rejected' | 'withdrawn' | 'cancelled'
export type Scope = 'mine' | 'team'
export type Filter = 'all' | Exclude<Status, 'withdrawn'>

export interface Session {
  username: string
  name: string
  role_tags: string[]
  approver_name: string
  can_review: boolean
  initial_password_pending?: boolean
}
export interface LeaveInput {
  start_date: string
  end_date: string
  start_slot: Slot
  end_slot: Slot
  note: string
}
export interface LeaveRequest extends LeaveInput {
  id: number
  employee_username: string | null
  employee_name: string
  status: Status
  approver_name: string
  created_at: string
  updated_at: string
  can_review: boolean
  cancellation_status: '' | 'pending' | 'approved' | 'rejected'
  can_request_cancellation: boolean
  can_review_cancellation: boolean
  can_withdraw: boolean
}
export interface LeaveList {
  items: LeaveRequest[]
  total: number
  page: number
  page_size: number
  counts: { total: number; pending: number; approved: number; rejected: number }
}
export interface LeaveEvent {
  id: number
  actor_name: string
  actor_username: string
  action: 'submitted' | 'decided' | 'withdrawn' | 'cancel_requested' | 'cancel_approved' | 'cancel_rejected' | 'attendance_synced'
  previous_status: Status | null
  new_status: Status
  comment: string
  created_at: string
}
export interface LeaveDetail { request: LeaveRequest; history: LeaveEvent[] }

export interface LeaveAttendancePreview {
  operation: 'approve' | 'cancel'
  fingerprint: string
  created_sheets: number
  updated_cells: number
  linked_cells: number
  skipped_cells: number
  can_apply: boolean
  needs_confirmation: boolean
  issues: { code: string; date: string; period: string; status: string; blocking: boolean }[]
}
