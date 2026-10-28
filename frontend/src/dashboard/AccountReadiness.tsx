import { useState } from 'react'
import { Box, Chip, Typography } from '@mui/material'
import { useWorkCenter } from './useWorkCenter'
import { ReadinessStatus } from './ReadinessStatus'

export function AccountReadiness({ token, viewer, username }: { token: string; viewer: string; username: string }) {
  const [today] = useState(() => new Date())
  const query = useWorkCenter(token, viewer, today.getMonth() + 1, today.getFullYear())
  const person = query.data?.people.find(item => item.username === username)
  if (!person) return query.error ? <Typography variant="caption" color="warning.main">État du dossier indisponible.</Typography> : null
  return <Box sx={{ display: 'flex', alignItems: 'center', flexWrap: 'wrap', gap: 1, mb: 1.5 }}>
    <Chip size="small" variant="outlined" color={person.attendance_ready ? 'success' : 'warning'} label={person.attendance_ready ? 'Contrat présent ce mois' : 'Contrat à vérifier ce mois'} />
    <ReadinessStatus person={person} />
  </Box>
}
