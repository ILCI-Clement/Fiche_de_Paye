import type { EmploymentInput, PersonnelInput } from './api'

export function missingRegistrationFields(value: PersonnelInput, period: EmploymentInput) {
  const fields: [keyof PersonnelInput, string][] = [['last_name', 'Nom'], ['first_name', 'Prénoms'], ['nationality', 'Nationalité'],
    ['sex', 'Sexe'], ['birth_date', 'Date de naissance'], ['address', 'Adresse'], ['job_title', 'Emploi'],
    ['qualification', 'Qualification'], ['personnel_type', 'Type de personnel']]
  const missing = fields.filter(([key]) => !value[key]?.trim()).map(([, label]) => label)
  if (!period.entry_date) missing.push('Date d’entrée')
  return missing
}

export function validateEmployment(value: EmploymentInput) {
  if (value.entry_date && value.exit_date && value.exit_date < value.entry_date) return 'La fin prévue du contrat ne peut pas précéder l’entrée.'
  if (value.entry_date && value.actual_exit_date && value.actual_exit_date < value.entry_date) return 'La sortie effective ne peut pas précéder l’entrée.'
  if (value.is_foreign_worker && (!value.work_permit_type.trim() || !value.work_permit_number.trim())) return 'Renseignez le type et le numéro du document autorisant le travail.'
  if (value.is_temporary_worker && (!value.contract_type || (value.contract_type === 'autre' && !value.other_contract.trim()) || !value.temporary_company_name.trim() || !value.temporary_company_address.trim())) return 'Renseignez le contrat, le nom et l’adresse de l’entreprise de travail temporaire.'
  if (value.is_young_worker && !value.youth_contract_type) return 'Sélectionnez le contrat du jeune travailleur ou de l’apprenti.'
  return ''
}
