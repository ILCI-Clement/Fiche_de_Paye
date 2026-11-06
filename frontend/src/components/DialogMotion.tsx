import { forwardRef } from 'react'
import { Grow, useMediaQuery } from '@mui/material'
import type { GrowProps } from '@mui/material/Grow'

/** Keep MUI's focus, mounting and close lifecycle while adding a short zoom. */
export const DialogMotion = forwardRef<unknown, GrowProps>(function DialogMotion(props, ref) {
  const reduced = useMediaQuery('(prefers-reduced-motion: reduce)')
  return <Grow {...props} ref={ref} timeout={reduced ? 0 : { enter: 220, exit: 160 }} />
})
