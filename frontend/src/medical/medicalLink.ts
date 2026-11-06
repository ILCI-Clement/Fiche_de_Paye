export function linkedMedical(search: string): number | null {
  const value = new URLSearchParams(search).get('medical')
  if (!value || !/^[1-9]\d{0,14}$/.test(value)) return null
  const id = Number(value)
  return Number.isSafeInteger(id) ? id : null
}
