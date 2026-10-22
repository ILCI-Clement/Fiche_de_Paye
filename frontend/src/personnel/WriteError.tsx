import { Alert, Box, Typography } from '@mui/material'

export function WriteError({ error, conflict }: { error: string; conflict: boolean }) {
  if (!error) return null
  return <Alert severity={conflict ? 'warning' : 'error'} sx={{ mb: 2 }}>
    <Box>{error}</Box>
    {conflict && <Typography variant="body2" sx={{ mt: 1 }}>Vos saisies sont conservées. Fermez ce formulaire puis actualisez la fiche pour vérifier les données récentes avant de recommencer. Aucune donnée distante n’a été remplacée automatiquement.</Typography>}
  </Alert>
}
