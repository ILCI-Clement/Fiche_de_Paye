export type Slot = 'full' | 'morning' | 'afternoon'
export type Status = 'pending' | 'approved' | 'rejected' | 'withdrawn'
export type Scope = 'mine' | 'team'
export type Filter = 'all' | Exclude<Status, 'withdrawn'>

export interface Session {
  username: string
  name: string
  role_tags: string[]
  approver_name: string
  can_review: boolean
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
  action: 'submitted' | 'decided' | 'withdrawn'
  previous_status: Status | null
  new_status: Status
  comment: string
  created_at: string
}
export interface LeaveDetail { request: LeaveRequest; history: LeaveEvent[] }
