export interface PersonName {
  full_name?: string | null
  display_name?: string | null
  email?: string | null
}

/** Login identifiers are deliberately excluded from display-name fallbacks. */
export function personDisplayName(person?: PersonName | null): string {
  const email = person?.email?.trim() || ''
  const prefix = /^[^@\s]+@[^@\s]+$/.test(email) ? email.slice(0, email.indexOf('@')) : ''
  return person?.full_name?.trim() || person?.display_name?.trim() || prefix || 'Nom non renseigné'
}

/** Distinguish namesakes in selectors without exposing the login identifier. */
export function personOptionLabel(person?: PersonName | null): string {
  const name = personDisplayName(person)
  const email = person?.email?.trim() || ''
  return /^[^@\s]+@[^@\s]+$/.test(email) ? `${name} · ${email}` : name
}
