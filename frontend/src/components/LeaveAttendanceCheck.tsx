import { Alert, Box, Checkbox, FormControlLabel, Typography } from '@mui/material'
import type { LeaveAttendancePreview } from '../types'
import { dateLabel } from '../leaveFormat'

const labels: Record<string, string> = {
  missing_contract: 'Contrat à compléter pour cette date',
  protected: 'Fiche signée, en cours de signature ou archivée : modification interdite',
  modified: 'Présence modifiée depuis l’approbation : vérification manuelle nécessaire',
  missing_sheet: 'Fiche d’origine absente : vérification manuelle nécessaire',
  conflict: 'Présence existante à remplacer par un congé payé',
}

export function LeaveAttendanceCheck({ data, accepted, onAccepted }: {
  data: LeaveAttendancePreview; accepted: boolean; onAccepted: (value: boolean) => void
}) {
  const issues = [...new Map(data.issues.map(issue => [
    `${issue.code}:${issue.date}:${issue.period}:${issue.status}`, issue,
  ])).values()]
  return <Box sx={{ mt: 2 }}>
    <Typography variant="body2" color="text.secondary">
      {data.created_sheets > 0 && `${data.created_sheets} fiche(s) à créer · `}
      {data.updated_cells} demi-journée(s) à {data.operation === 'cancel' ? 'restaurer' : 'mettre à jour'}
      {data.linked_cells > 0 && ` · ${data.linked_cells} déjà en congé payé`}
    </Typography>
    {issues.length > 0 && <Alert severity={data.can_apply ? 'warning' : 'error'} sx={{ mt: 1 }}>
      <Box sx={{ maxHeight: 160, overflowY: 'auto' }}>
        {issues.map((issue, i) => <Typography key={i} variant="body2" sx={{ mb: .5 }}>
          {dateLabel(issue.date)} {issue.period === 'matin' ? 'matin' : issue.period === 'aprem' ? 'après-midi' : ''}
          {' · '}{labels[issue.code] || 'Vérification nécessaire'}{issue.status && ` (${issue.status})`}
        </Typography>)}
      </Box>
    </Alert>}
    {data.can_apply && data.needs_confirmation && <FormControlLabel sx={{ mt: 1 }}
      control={<Checkbox checked={accepted} onChange={event => onAccepted(event.target.checked)} />}
      label={<Typography variant="body2">Je confirme le remplacement des présences indiquées ci-dessus.</Typography>} />}
  </Box>
}
