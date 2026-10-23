import type { EmploymentInput, PersonnelInput } from './api'

export const personnelFieldLabels = {
  full_name: 'Nom et prénoms', last_name: 'Nom', first_name: 'Prénoms', personnel_type: 'Type de personnel',
  nationality: 'Nationalité', sex: 'Sexe', birth_date: 'Date de naissance', address: 'Adresse', job_title: 'Emploi', qualification: 'Qualification',
}

export function missingRegistrationFields(value: PersonnelInput, period: EmploymentInput) {
  const keys = Object.keys(personnelFieldLabels) as (keyof typeof personnelFieldLabels)[]
  const missing = keys.filter(key => key !== 'full_name' && !value[key]?.trim()).map(key => personnelFieldLabels[key])
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
