export function completionSection(code: string) {
  if (code === 'manager') return 'organization'
  if (code === 'contract') return 'contract'
  return 'register'
}

export function missingPersonnelField(person: Record<string, unknown>, code = '') {
  if (code === 'email') return 'email'
  return ['last_name', 'first_name', 'personnel_type', 'nationality', 'sex', 'birth_date', 'address', 'job_title', 'qualification']
    .find(key => !String(person[key] ?? '').trim())
}
