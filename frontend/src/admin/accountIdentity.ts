/** Human labels may change or repeat; destructive confirmation remains account-specific. */
export function accountConfirmation(account: { username: string; email?: string | null }) {
  const email = account.email?.trim()
  return email ? { expected: email, label: 'Adresse e-mail de confirmation' }
    : { expected: account.username, label: 'Identifiant technique de confirmation' }
}
