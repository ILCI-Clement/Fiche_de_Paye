import type { SavedConfig, Workspace } from './types'

export function canAdoptWorkspaceRevision(previous: string, incoming: string, state: { dirty: boolean; busy: boolean; pending: boolean }) {
  return previous !== incoming && !state.dirty && !state.busy && !state.pending
}

/** Finish canonical synchronization before an editable workspace is shown. No retry on conflicts. */
export async function refreshWorkspaceContracts(workspace: Workspace, refresh: (revision: string) => Promise<SavedConfig>): Promise<Workspace> {
  if (!workspace.contract_refresh_required) return workspace
  const saved = await refresh(workspace.revision)
  return { ...workspace, ...saved, contract_refresh_required: false }
}
