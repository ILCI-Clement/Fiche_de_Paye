import { useState } from 'react'
import { Box, Button, Chip, Dialog, DialogActions, DialogContent, DialogTitle, Divider, Typography } from '@mui/material'
import type { PersonReadiness } from './workCenterApi'
import { useWorkNavigation } from '../workspaceNavigation'

export function ReadinessStatus({ person }: { person: PersonReadiness }) {
  const [open, setOpen] = useState(false)
  const navigation = useWorkNavigation()
  return <>
    <Button size="small" color={person.issues.length ? 'warning' : 'success'} onClick={() => setOpen(true)} aria-label={`État du dossier de ${person.display_name}`} sx={{ whiteSpace: 'nowrap', fontSize: 12 }}>
      {person.issues.length ? `${person.issues.length} point(s) à vérifier` : 'Dossier prêt'}
    </Button>
    <Dialog open={open} onClose={() => setOpen(false)} fullWidth maxWidth="sm">
      <DialogTitle>État du dossier · {person.display_name}</DialogTitle>
      <DialogContent>
        <Box sx={{ display: 'flex', gap: 1, flexWrap: 'wrap', mb: 2 }}>
          <Chip size="small" variant="outlined" color={person.attendance_ready ? 'success' : 'warning'} label={person.attendance_ready ? 'Contrat présent pour ce mois' : 'Contrat à vérifier pour ce mois'} />
          <Chip size="small" variant="outlined" color={person.contacts_ready ? 'success' : 'warning'} label={person.contacts_ready ? 'Contacts renseignés' : 'Contacts à compléter'} />
        </Box>
        {!person.issues.length && <Typography variant="body2">Les prérequis du dossier sont renseignés. Chaque fiche reste contrôlée avant envoi.</Typography>}
        {person.issues.map(issue => <Box key={issue.code} sx={{ mb: 1 }}><Box sx={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 1, py: .75 }}>
          <Box><Typography variant="body2">{issue.label}</Typography>{issue.domain === 'register' && <Typography variant="caption" color="text.secondary">Sans blocage de la création d’une fiche de temps.</Typography>}
          {!issue.destination && <Typography variant="caption" color="text.secondary" sx={{ display: 'block' }}>À faire vérifier par un administrateur ou le responsable direct.</Typography>}</Box>
          {issue.destination && navigation && <Button size="small" sx={{ flexShrink: 0 }} onClick={() => { setOpen(false); navigation.navigate(issue.destination!) }}>Ouvrir</Button>}
        </Box><Divider /></Box>)}
      </DialogContent>
      <DialogActions><Button onClick={() => setOpen(false)}>Fermer</Button></DialogActions>
    </Dialog>
  </>
}
