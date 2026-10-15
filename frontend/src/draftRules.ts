export type Draft = { dirty: boolean; busy: boolean; label: string }

export function draftDecision(drafts: Record<string, Draft>, only?: string, exceptLabel?: string) {
  const relevant = only ? (drafts[only] ? [drafts[only]] : []) : Object.values(drafts).filter(draft => draft.label !== exceptLabel)
  return {
    blocked: relevant.some(draft => draft.busy),
    labels: [...new Set(relevant.filter(draft => draft.dirty).map(draft => draft.label))],
  }
}
