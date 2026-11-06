// Business functions and application access share the existing API role_tags field.
export const businessFunction = (roles: readonly string[]) => roles.includes('Responsable') ? 'Responsable' : roles.includes('Employe') ? 'Employe' : ''

export function withBusinessFunction(roles: readonly string[], selected: string): string[] {
  return [...roles.filter(role => role !== 'Employe' && role !== 'Responsable'),
    ...(['Responsable', 'Employe'].includes(selected) ? [selected] : [])]
}

export function withSystemAccess(roles: readonly string[], access: 'Admin' | 'SuperAdmin', checked: boolean): string[] {
  const business = roles.filter(role => role !== 'Admin' && role !== 'SuperAdmin')
  if (access === 'SuperAdmin') return [...business, ...(checked ? ['SuperAdmin', 'Admin'] : roles.includes('Admin') ? ['Admin'] : [])]
  return [...business, ...(roles.includes('SuperAdmin') ? ['SuperAdmin', 'Admin'] : checked ? ['Admin'] : [])]
}
