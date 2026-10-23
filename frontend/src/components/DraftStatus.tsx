import type { ReactNode } from 'react'
import { Paper, Tooltip, Typography } from '@mui/material'
import { useWorkspaceGuard } from '../draftGuard'

/** A persistent, non-blocking hint; navigation confirmation stays in DraftGuardProvider. */
export function DraftStatus({ status, children }: { status?: string; children?: ReactNode }) {
  const guard = useWorkspaceGuard()
  const text = status || (guard?.dirtyLabels.length ? 'Non enregistré' : '')
  if (!text && !children) return null
  return <Paper className="workspace-draft-status" elevation={0}>
    {text && <Tooltip title={guard?.dirtyLabels.join(' · ') || text} enterDelay={650} arrow>
      <Typography component="span" role="status" aria-live="polite" aria-atomic="true" variant="caption" color="text.secondary">{text}</Typography>
    </Tooltip>}
    {children}
  </Paper>
}
