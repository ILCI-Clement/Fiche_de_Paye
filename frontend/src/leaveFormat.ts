import type { LeaveRequest, Slot, Status } from './types'

export const slotLabels: Record<Slot, string> = { full: 'Journée entière', morning: 'Matin', afternoon: 'Après-midi' }
export const statusLabels: Record<Status, string> = { pending: 'En attente', approved: 'Approuvée', rejected: 'Refusée', withdrawn: 'Retirée' }
export function dateLabel(value: string) {
  return new Intl.DateTimeFormat('fr-FR', { day: 'numeric', month: 'short', year: 'numeric' }).format(new Date(`${value}T12:00:00`))
}
export function periodLabel(item: LeaveRequest) {
  return item.start_date === item.end_date ? dateLabel(item.start_date) : `${dateLabel(item.start_date)} → ${dateLabel(item.end_date)}`
}
export function slotLabel(item: LeaveRequest) {
  return item.start_date === item.end_date ? slotLabels[item.start_slot]
    : `${slotLabels[item.start_slot]} au début · ${slotLabels[item.end_slot]} à la fin`
}
export function timeLabel(value: string) {
  return new Intl.DateTimeFormat('fr-FR', { dateStyle: 'medium', timeStyle: 'short' }).format(new Date(value))
}
export function initials(name: string) { return name.split(/[ ._-]+/).filter(Boolean).slice(0, 2).map(x => x[0]).join('').toUpperCase() }
