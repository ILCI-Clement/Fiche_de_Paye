/** A link selects a request; it never carries credentials or performs a decision. */
export function linkedLeave(search: string): number | null {
  const value = new URLSearchParams(search).get('leave')
  if (!value || !/^[1-9]\d{0,14}$/.test(value)) return null
  const id = Number(value)
  return Number.isSafeInteger(id) ? id : null
}
