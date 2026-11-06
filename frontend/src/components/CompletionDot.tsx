import { Box, Tooltip } from '@mui/material'

/** A section warning, without exposing or duplicating personal field values. */
export function CompletionDot({ issues }: { issues: string[] }) {
  if (!issues.length) return null
  const label = `À compléter : ${issues.join(' · ')}`
  return <Tooltip title={label} arrow><Box component="span" role="img" aria-label={label} tabIndex={0}
    sx={{ display: 'inline-block', width: 8, height: 8, borderRadius: '50%', bgcolor: '#F4B400', flexShrink: 0, mr: 1.25, verticalAlign: 'middle' }} /></Tooltip>
}
