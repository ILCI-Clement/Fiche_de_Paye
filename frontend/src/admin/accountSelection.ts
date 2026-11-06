/** An explicit navigation target must never fall back to somebody else's record. */
export function selectedAccount<T extends { username: string }>(accounts: readonly T[], selected: string): T | undefined {
  return selected ? accounts.find(account => account.username === selected) : accounts[0]
}
