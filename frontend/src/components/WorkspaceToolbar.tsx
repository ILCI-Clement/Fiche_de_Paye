import type { ReactNode } from 'react'
import { createPortal } from 'react-dom'
import { Box } from '@mui/material'
import { useWorkspaceToolbar } from '../workspaceToolbar'

/** Keep form state in its page; only the active page may render controls in the shell. */
export function WorkspaceToolbar({ children, placement = 'tools' }: { children: ReactNode; placement?: 'tools' | 'inline' }) {
  const slot = useWorkspaceToolbar()
  const toolbar = <Box className={`workspace-action-bar${placement === 'inline' ? ' workspace-action-bar--inline' : ''}`} role="toolbar" aria-label="Actions de la page">{children}</Box>
  // Standalone previews can display their tools without an application shell.
  if (!slot) return toolbar
  const target = placement === 'inline' ? slot.inlineTarget : slot.target
  if (!slot.active || !target) return null
  return createPortal(toolbar, target)
}
