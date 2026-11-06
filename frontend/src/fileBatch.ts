export type FileAction = 'archive' | 'restore' | 'delete'
export interface FileBatchResult { count: number; cleanup_pending: boolean }

export const fileActionLabels: Record<FileAction, string> = {
  archive: 'Archiver', restore: 'Restaurer', delete: 'Supprimer définitivement',
}
