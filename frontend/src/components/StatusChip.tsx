import { Chip } from '@mui/material'
import type { Status } from '../types'
import { statusLabels } from '../leaveFormat'
import { alpha } from '@mui/material/styles'

const colors = { pending: 'warning', approved: 'success', rejected: 'error', withdrawn: 'secondary' } as const

export function StatusChip({ status }: { status: Status }) {
  return <Chip label={statusLabels[status]} size="small" sx={theme => ({
    bgcolor: alpha(theme.palette[colors[status]].main, .14), color: theme.palette[colors[status]].main,
  })} />
}
