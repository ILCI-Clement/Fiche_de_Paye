export interface Schedule { actif: boolean; m1: string; m2: string; a1: string; a2: string }
export interface Fiche {
  id: string | number
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
export interface Account { username: string; email: string; contracts: unknown[]; group_names?: string[] }
export interface SavedConfig { config: Config; revision: string }
export interface Workspace extends SavedConfig { accounts: Account[]; can_edit_identity: boolean }
export interface Day { date: string; number: number; weekday: number; disabled: boolean; matin: [string, string]; aprem: [string, string] }
export interface Cell { date: string; period: 'matin' | 'aprem' }

export function ficheLabel(fiche: Fiche) {
  const name = fiche.type === 'Stagiaire' ? `${fiche.prenom_stagiaire || ''} ${fiche.nom_stagiaire || ''}`.trim() : fiche.nom
  const label = (name || fiche.email_employe || fiche.account_username || `Fiche ${fiche.id}`).split('@')[0]
  return fiche.contract_id ? `${label} · ${fiche.sheet_period_start?.slice(8)}–${fiche.sheet_period_end?.slice(8)}` : label
}

export function accountOptionLabel(account: Account) {
  return `${account.group_names?.join(', ') || 'Unknown'} · ${account.email.split('@')[0]} (${account.username.split('@')[0]})`
}

