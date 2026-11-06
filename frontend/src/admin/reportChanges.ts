export function reportChanges(accounts: { username: string; manager_id: string | null }[], manager: string, selected: readonly string[]) {
  const chosen = new Set(selected)
  return accounts.filter(person => (person.manager_id === manager) !== chosen.has(person.username))
    .map(person => ({ username: person.username, expected_manager_id: person.manager_id, assigned: chosen.has(person.username) }))
}
