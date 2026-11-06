import { Chip } from '@mui/material'
import type { Status } from '../types'
import { statusLabels } from '../leaveFormat'
import { alpha } from '@mui/material/styles'

const colors = { pending: 'warning', approved: 'success', rejected: 'error', withdrawn: 'secondary', cancelled: 'secondary' } as const

export function StatusChip({ status, cancellationPending = false }: { status: Status; cancellationPending?: boolean }) {
  return <Chip label={cancellationPending ? 'Annulation en attente' : statusLabels[status]} size="small" sx={theme => ({
    bgcolor: alpha(theme.palette[colors[status]].main, .14), color: theme.palette[colors[status]].main,
  })} />
}
