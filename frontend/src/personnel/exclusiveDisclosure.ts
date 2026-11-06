/** Coordinates visibility only; never stores a sensitive value or an employee ID. */
export function createExclusiveDisclosure() {
  let current: { revoke: () => void } | null = null
  return {
    claim(revoke: () => void) {
      const previous = current
      const owner = { revoke }
      current = owner
      // Revoke before the next request starts, including requests still in flight.
      previous?.revoke()
      return {
        isCurrent: () => current === owner,
        release: () => { if (current === owner) current = null },
      }
    },
  }
}

export const sensitiveNumberDisclosure = createExclusiveDisclosure()
