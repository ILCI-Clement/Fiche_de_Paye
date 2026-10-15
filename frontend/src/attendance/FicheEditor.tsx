import { Accordion, AccordionDetails, AccordionSummary, Box, Checkbox, FormControlLabel, MenuItem, TextField, Typography } from '@mui/material'
import ExpandMoreRounded from '@mui/icons-material/ExpandMoreRounded'
import type { Fiche, Schedule } from './types'

const weekdays = ['Lundi', 'Mardi', 'Mercredi', 'Jeudi', 'Vendredi', 'Samedi', 'Dimanche']

export function FicheEditor({ fiche, admin, disabled, onChange }: { fiche: Fiche; admin: boolean; disabled: boolean; onChange: (data: Partial<Fiche>) => void }) {
  const intern = fiche.type === 'Stagiaire'
  const locked = disabled || !admin
  const contractLocked = locked || Boolean(fiche.contract_id)
  function field(key: string, label: string, options?: { date?: boolean; numeric?: boolean; readOnly?: boolean }) {
    return <TextField key={key} label={label} type={options?.date ? 'date' : options?.numeric ? 'number' : 'text'} fullWidth
      value={String(fiche[key] ?? '') === 'Pas de fin' && options?.date ? '' : String(fiche[key] ?? '')}
      disabled={Boolean(options?.readOnly)} slotProps={{ inputLabel: options?.date ? { shrink: true } : undefined,
        htmlInput: options?.numeric ? { min: 0, step: 'any' } : undefined }}
      onChange={e => onChange({ [key]: options?.numeric ? Number(e.target.value) : e.target.value || null })} />
  }
  function updateSchedule(day: string, values: Partial<Schedule>) {
    const schedule = fiche.planning_detail?.[day] || { actif: !['Samedi', 'Dimanche'].includes(day), m1: '09:00', m2: '12:00', a1: '13:00', a2: '17:00' }
    onChange({ planning_detail: { ...fiche.planning_detail, [day]: { ...schedule, ...values } } })
  }
  return <Box className="fiche-details">
    <Accordion disableGutters><AccordionSummary expandIcon={<ExpandMoreRounded />}><Typography sx={{ fontWeight: 650 }}>Informations personnelles et contrat</Typography></AccordionSummary>
      <AccordionDetails><Box className="fiche-fields">
        {field(intern ? 'nom_stagiaire' : 'nom', intern ? 'Nom du stagiaire' : 'NOM Prénom (Employé)', { readOnly: locked })}
        {intern && field('prenom_stagiaire', 'Prénom du stagiaire', { readOnly: locked })}
        {field('email_employe', 'E-mail de l’employé', { readOnly: locked })}
        {field('responsable', 'NOM Prénom (Responsable)', { readOnly: locked })}
        {field('email_responsable', 'E-mail du responsable', { readOnly: locked })}
        <TextField select label="Type de contrat" value={fiche.type || 'Salarié'} disabled={contractLocked} onChange={e => onChange({ type: e.target.value })}>
          <MenuItem value="Salarié">Salarié</MenuItem><MenuItem value="Stagiaire">Stagiaire</MenuItem>
        </TextField>
        {field(intern ? 'dds' : 'ddc', intern ? 'Début de stage' : 'Début de contrat', { date: true, readOnly: contractLocked })}
        {!(fiche.cdi && !intern) && field(intern ? 'fds' : 'fdc', intern ? 'Fin de stage' : 'Fin de contrat', { date: true, readOnly: contractLocked })}
        {!intern && <FormControlLabel label="Contrat CDI" control={<Checkbox checked={Boolean(fiche.cdi)} disabled={contractLocked} onChange={e => onChange({ cdi: e.target.checked, fdc: e.target.checked ? 'Pas de fin' : null })} />} />}
      </Box>
      {fiche.contract_id && <Typography variant="caption" color="text.secondary" sx={{ display: 'block', mt: 2 }}>Contrat #{fiche.contract_id} · Période de fiche : {fiche.sheet_period_start} → {fiche.sheet_period_end}</Typography>}
      </AccordionDetails>
    </Accordion>
    {!intern && <Accordion disableGutters><AccordionSummary expandIcon={<ExpandMoreRounded />}><Typography sx={{ fontWeight: 650 }}>Temps partiel / Planning hebdomadaire</Typography></AccordionSummary>
      <AccordionDetails>{weekdays.map(day => {
        const schedule = fiche.planning_detail?.[day] || { actif: !['Samedi', 'Dimanche'].includes(day), m1: '09:00', m2: '12:00', a1: '13:00', a2: '17:00' }
        return <Box key={day} className="schedule-row">
          <FormControlLabel label={day} control={<Checkbox checked={schedule.actif} disabled={disabled} onChange={e => updateSchedule(day, { actif: e.target.checked })} />} />
          {(['m1', 'm2', 'a1', 'a2'] as const).map((key, i) => <TextField key={key} type="time" label={['Matin de', 'à', 'Après-midi de', 'à'][i]}
            value={schedule[key]} disabled={disabled || !schedule.actif} slotProps={{ inputLabel: { shrink: true } }} onChange={e => updateSchedule(day, { [key]: e.target.value })} />)}
        </Box>
      })}</AccordionDetails>
    </Accordion>}
    {intern && <Accordion disableGutters><AccordionSummary expandIcon={<ExpandMoreRounded />}><Typography sx={{ fontWeight: 650 }}>Rémunération et présence du stage</Typography></AccordionSummary>
      <AccordionDetails><Box className="fiche-fields">
        {field('nb_jours', 'Nombre de jours', { numeric: true, readOnly: disabled })}
        {field('nb_heures_jour', 'Heures par jour', { numeric: true, readOnly: disabled })}
        {field('taux_horaire', 'Taux horaire (€)', { numeric: true, readOnly: locked })}
        {field('facture_mensuelle', 'Facture mensuelle (€)', { numeric: true, readOnly: locked })}
        {field('taux', 'Taux (%)', { numeric: true, readOnly: locked })}
        {field('transport', 'Transport', { readOnly: locked })}
      </Box></AccordionDetails>
    </Accordion>}
  </Box>
}
