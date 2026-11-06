import { useState } from 'react'
import { Box, Typography } from '@mui/material'
import { useWorkCenter } from './useWorkCenter'
import { ReadinessStatus } from './ReadinessStatus'

export function AccountReadiness({ token, viewer, username }: { token: string; viewer: string; username: string }) {
  const [today] = useState(() => new Date())
  const query = useWorkCenter(token, viewer, today.getMonth() + 1, today.getFullYear())
  const person = query.data?.people.find(item => item.username === username)
  if (person && !person.issues.length) return null
  if (!person) return query.error ? <Typography variant="caption" color="warning.main">État du dossier indisponible.</Typography> : null
  return <Box sx={{ display: 'flex', alignItems: 'center', flexWrap: 'wrap', gap: 1, mb: 1.5 }}>
    <ReadinessStatus person={person} />
  </Box>
}
