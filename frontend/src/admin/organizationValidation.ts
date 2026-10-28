import type { OrganizationInput } from './api'

export function organizationValidation(value: OrganizationInput): string {
  const roles = value.role_tags
  if (!roles.length) return 'Choisissez au moins une fonction.'
  if (roles.includes('Employe') && !value.manager_id) return 'Choisissez un responsable direct pour cet employé.'
  if (roles.includes('Responsable') && !roles.includes('Admin') && !roles.includes('SuperAdmin') && !value.managed_group_ids.length) {
    return 'Choisissez au moins un département géré pour ce Responsable.'
  }
  if (value.managed_group_ids.length && !roles.includes('Responsable')) return 'Choisissez la fonction Responsable pour gérer des départements.'
  return ''
}
