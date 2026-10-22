import { personDisplayName } from '../personDisplayName.ts'

export interface Schedule { actif: boolean; m1: string; m2: string; a1: string; a2: string }
export interface Fiche {
  id: string | number
  is_archived?: boolean
  type?: string
  nom?: string
  nom_stagiaire?: string
  prenom_stagiaire?: string
  account_username?: string
  email_employe?: string
  email_responsable?: string
  responsable?: string
  ddc?: string | null
  fdc?: string | null
  dds?: string | null
  fds?: string | null
  cdi?: boolean
  planning_detail?: Record<string, Schedule>
  nb_jours?: number
  nb_heures_jour?: number
  taux_horaire?: number
  facture_mensuelle?: number
  taux?: number
  transport?: string
  contract_id?: number
  sheet_month?: number
  sheet_year?: number
  sheet_period_start?: string
  sheet_period_end?: string
  [key: string]: unknown
}
export interface Config { mois: number; annee: number; employes_data: Fiche[]; [key: string]: unknown }
export interface Account { username: string; email: string; full_name?: string | null; display_name?: string; manager_display_name?: string; contracts: unknown[]; group_names?: string[] }
export interface SavedConfig { config: Config; revision: string; person_labels?: Record<string, string> }
export interface Workspace extends SavedConfig { accounts: Account[]; can_edit_identity: boolean; can_delete_sheets: boolean; can_archive_sheets: boolean }
export interface Day { date: string; number: number; weekday: number; disabled: boolean; matin: [string, string]; aprem: [string, string] }
export interface Cell { date: string; period: 'matin' | 'aprem' }

/** Labels are presentation metadata: never merge them into a saved or dirty fiche. */
export function attendancePersonName(fiche: Fiche, workspace: Pick<Workspace, 'accounts' | 'person_labels'>) {
  const account = workspace.accounts.find(person => person.username === fiche.account_username)
  return workspace.person_labels?.[String(fiche.id)] || personDisplayName(account || { email: fiche.email_employe })
}

export function withAttendanceLabels(workspace: Workspace, saved: Pick<SavedConfig, 'person_labels'>): Workspace {
  return saved.person_labels ? { ...workspace, person_labels: saved.person_labels } : workspace
}

/** Existing periods come only from explicit account links and saved month bindings. */
export function accountSheetPeriods(fiches: Fiche[], username: string, archived = false) {
  const periods = new Map<string, { month: number; year: number; id: string }>()
  for (const fiche of fiches) {
    if (!username || fiche.account_username !== username || !!fiche.is_archived !== archived) continue
    const month = fiche.sheet_month, year = fiche.sheet_year
    if (typeof month !== 'number' || typeof year !== 'number' || !Number.isInteger(month) || !Number.isInteger(year) || month < 1 || month > 12 || year < 2000 || year > 2100) continue
    const key = `${year}-${month}`
    if (!periods.has(key)) periods.set(key, { month, year, id: String(fiche.id) })
  }
  return [...periods.values()].sort((a, b) => b.year - a.year || b.month - a.month)
}

export function ficheLabel(fiche: Fiche, displayName?: string) {
  const label = personDisplayName({ display_name: displayName, email: fiche.email_employe })
  return fiche.contract_id ? `${label} · ${fiche.sheet_period_start?.slice(8)}–${fiche.sheet_period_end?.slice(8)}` : label
}

export function accountOptionLabel(account: Account) {
  return `${account.group_names?.join(', ') || 'Unknown'} · ${personDisplayName(account)}`
}

