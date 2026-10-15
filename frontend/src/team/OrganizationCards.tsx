import { Accordion, AccordionDetails, AccordionSummary, Alert, Box, Chip, Paper, Typography } from '@mui/material'
import ExpandMoreRounded from '@mui/icons-material/ExpandMoreRounded'
import type { Department, Organization } from './api'

function DepartmentCard({ department, general = false }: { department: Department; general?: boolean }) {
  return <Paper variant="outlined" sx={{ p: 2.5, minWidth: 0, ...(general ? { maxWidth: 440, mx: 'auto', borderColor: 'primary.main' } : {}) }}>
    <Typography variant="h6" sx={{ overflowWrap: 'anywhere', borderBottom: '1px solid', borderColor: 'divider', pb: 1.5 }}>{department.name}</Typography>
    <Typography variant="overline" color="text.secondary" sx={{ display: 'block', mt: 1 }}>Responsable</Typography>
    <Typography color="primary.main" sx={{ fontWeight: 650, overflowWrap: 'anywhere' }}>{department.leaders.join(', ') || 'Responsable à désigner'}</Typography>
    <Typography variant="overline" color="text.secondary" sx={{ display: 'block', mt: 1 }}>Équipe</Typography>
    {department.members.length ? <Box component="ul" sx={{ listStyle: 'none', p: 0, m: 0 }}>{department.members.map((name, index) => <Typography component="li" key={`${name}-${index}`} variant="body2" sx={{ overflowWrap: 'anywhere', py: .4 }}>{name}</Typography>)}</Box>
      : <Typography variant="body2" color="text.secondary">Aucun membre affecté</Typography>}
  </Paper>
}

export function OrganizationCards({ model }: { model: Organization }) {
  const general = model.departments.find(d => d.id === model.general_id)
  return <Box>
      <Box sx={{ display: 'flex', flexWrap: 'wrap', gap: 1, mb: 2 }}><Chip size="small" variant="outlined" label={`${model.departments.length} départements actifs`} /><Chip size="small" variant="outlined" label={`${model.people_count} personnes`} /></Box>
      {general && <DepartmentCard department={general} general />}
      <Box sx={{ display: 'grid', gridTemplateColumns: { xs: '1fr', md: 'repeat(2, minmax(0, 1fr))', xl: 'repeat(3, minmax(0, 1fr))' }, gap: 2, mt: general ? 2 : 0 }}>
        {model.departments.filter(d => d.id !== model.general_id).map(department => <DepartmentCard key={department.id} department={department} />)}
      </Box>
      {!model.departments.length && <Typography color="text.secondary">Aucun département actif.</Typography>}
      {!!model.unassigned.length && <Paper variant="outlined" sx={{ p: 2, mt: 2 }}><Typography variant="h6">Personnes sans département actif</Typography>
        {model.unassigned.map((person, index) => <Typography key={index} variant="body2" sx={{ mt: 1, overflowWrap: 'anywhere' }}>{person.Nom} · {person.Situation}</Typography>)}
      </Paper>}
      {!!model.issues.length && <Accordion disableGutters sx={{ mt: 2 }}><AccordionSummary expandIcon={<ExpandMoreRounded />}><Typography>Affectations à vérifier ({model.issues.length})</Typography></AccordionSummary>
        <AccordionDetails>{model.issues.map((issue, i) => <Alert key={i} severity="warning" sx={{ mb: 1 }}>{issue}</Alert>)}</AccordionDetails>
      </Accordion>}
  </Box>
}
