import { IconButton, Tooltip } from '@mui/material'
import DeleteOutlineRounded from '@mui/icons-material/DeleteOutlineRounded'

export function DeleteAction({ label, disabled = false, onClick }: { label: string; disabled?: boolean; onClick: () => void }) {
  return <Tooltip title={label} arrow describeChild>
    <span style={{ display: 'inline-flex' }}>
      <IconButton color="error" aria-label={label} disabled={disabled} onClick={onClick} sx={{ width: 44, height: 44 }}>
        <DeleteOutlineRounded fontSize="small" />
      </IconButton>
    </span>
  </Tooltip>
}
