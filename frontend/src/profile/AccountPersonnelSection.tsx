import { useState } from 'react'
import { useQuery, useQueryClient } from '@tanstack/react-query'
import { Accordion, AccordionDetails, AccordionSummary, Alert, Box, Button, Chip, CircularProgress, Dialog, DialogActions, DialogContent, DialogTitle, MenuItem, Paper, Table, TableBody, TableCell, TableHead, TableRow, TextField, Typography } from '@mui/material'
import ExpandMoreRounded from '@mui/icons-material/ExpandMoreRounded'
import LinkOutlined from '@mui/icons-material/LinkOutlined'
import { useDraftGuard } from '../draftGuard'
import { contractLabel, dateLabel, typeLabel } from '../personnel/api'
import { useRegistryAction } from '../personnel/useRegistryAction'
import { WriteError } from '../personnel/WriteError'
import { accountPersonnelApi } from './accountPersonnelApi'
import type { AccountPersonnel, PersonnelChoice } from './accountPersonnelApi'

/** Account-bound self-service data; only the server can enable association edits. */
export function AccountPersonnelSection({ token, username, viewer, disabled = false, embedded = false }: { token: string; username: string; viewer: string; disabled?: boolean; embedded?: boolean }) {
  const client = useQueryClient()
  const profile = useQuery({ queryKey: ['account-personnel', viewer, username], queryFn: ({ signal }) => accountPersonnelApi.get(token, username, signal), retry: false, gcTime: 0 })
  const [editing, setEditing] = useState(false)
  const [message, setMessage] = useState('')
  const data = profile.data, employee = data?.employee
  const content = <>
        {profile.isPending ? <CircularProgress size={24} /> : profile.error ? <Alert severity="error">{profile.error.message}<Button onClick={() => void profile.refetch()}>Réessayer</Button></Alert> : data && <>
          <Box sx={{ display: 'flex', alignItems: 'center', flexWrap: 'wrap', gap: 1, mb: 2 }}>
            {employee && <Chip size="small" variant="outlined" label={`Registre n° ${employee.registration_number}`} />}
            {employee?.is_archived && <Chip size="small" label="Fiche archivée" />}
            <Typography variant="body2" color="text.secondary" sx={{ flex: 1 }}>Informations en lecture seule, tenues à jour par un administrateur.</Typography>
            {data.can_manage_link && <Button startIcon={<LinkOutlined />} size="small" disabled={disabled} onClick={() => { setMessage(''); setEditing(true) }}>{employee ? 'Modifier l’association' : 'Associer une fiche'}</Button>}
          </Box>
          {message && <Alert severity="success" sx={{ mb: 2 }}>{message}</Alert>}
          {!employee ? <Typography color="text.secondary">Aucune fiche du registre n’est associée à ce compte.{!data.can_manage_link ? ' Contactez un administrateur pour compléter votre profil.' : ''}</Typography> : <>
            <Box sx={{ display: 'grid', gridTemplateColumns: { xs: '1fr', sm: '1fr 1fr' }, gap: 2 }}>
              {[
                ['Nom et prénoms', employee.full_name], ['Type de personnel', typeLabel(employee.personnel_type)],
                ['Nationalité', employee.nationality], ['Sexe', employee.sex === 'F' ? 'Femme' : employee.sex === 'M' ? 'Homme' : ''],
                ['Date de naissance', dateLabel(employee.birth_date)], ['Adresse', employee.address],
                ['Emploi', employee.job_title], ['Qualification', employee.qualification],
              ].map(([label, value]) => <Box key={label} sx={{ minWidth: 0 }}><Typography variant="caption" color="text.secondary">{label}</Typography><Typography variant="body2" sx={{ overflowWrap: 'anywhere', whiteSpace: 'pre-line' }}>{value || 'Non renseigné'}</Typography></Box>)}
            </Box>
            <Typography variant="subtitle2" sx={{ mt: 3, mb: 1 }}>Périodes d’emploi inscrites au registre</Typography>
            {!data.periods.length ? <Typography variant="body2" color="text.secondary">Aucune période renseignée.</Typography> : <Table size="small" sx={{ tableLayout: 'fixed', '& td, & th': { overflowWrap: 'anywhere', px: 1 } }} aria-label="Périodes d’emploi du profil">
              <TableHead><TableRow><TableCell>Date d’entrée</TableCell><TableCell>Date de sortie</TableCell><TableCell>Contrat</TableCell></TableRow></TableHead>
              <TableBody>{data.periods.map(period => <TableRow key={period.period_id}><TableCell>{dateLabel(period.entry_date)}</TableCell><TableCell>{dateLabel(period.exit_date)}</TableCell><TableCell>{contractLabel(period)}</TableCell></TableRow>)}</TableBody>
            </Table>}
            <Typography variant="caption" color="text.secondary" sx={{ display: 'block', mt: 1 }}>Ces dates du registre sont indépendantes des contrats utilisés pour les fiches de présence.</Typography>
          </>}
        </>}
  </>
  return <>
    {embedded ? <Paper sx={{ p: { xs: 2, md: 3 }, width: '100%' }}>{content}</Paper> : <Accordion disableGutters sx={{ mb: 2 }}>
      <AccordionSummary expandIcon={<ExpandMoreRounded />}><Typography>Données du registre du personnel</Typography></AccordionSummary>
      <AccordionDetails>{content}</AccordionDetails>
    </Accordion>}
    {editing && data?.can_manage_link && <AssociationDialog token={token} username={username} viewer={viewer} initial={data} onClose={() => setEditing(false)} onSaved={async result => {
      client.setQueryData(['account-personnel', viewer, username], result)
      setEditing(false); setMessage(result.employee ? 'Fiche du registre associée au compte.' : 'Association retirée. La fiche et son historique sont conservés.')
      await client.invalidateQueries()
    }} />}
  </>
}

function AssociationDialog({ token, username, viewer, initial, onClose, onSaved }: { token: string; username: string; viewer: string; initial: AccountPersonnel; onClose: () => void; onSaved: (result: AccountPersonnel) => Promise<void> }) {
  const [search, setSearch] = useState('')
  const [selected, setSelected] = useState<PersonnelChoice | null>(() => initial.employee)
  const choices = useQuery({ queryKey: ['account-personnel-choices', viewer, username, search], queryFn: ({ signal }) => accountPersonnelApi.choices(token, username, search, signal), retry: false, gcTime: 0 })
  const action = useRegistryAction()
  const changed = (selected?.employee_id ?? null) !== (initial.employee?.employee_id ?? null)
  const close = useDraftGuard(changed, action.busy, 'Association au registre du personnel')
  const options = [...(choices.data?.employees || [])]
  if (selected && !options.some(item => item.employee_id === selected.employee_id)) options.unshift(selected)
  async function save() {
    if (!changed || action.busy) return
    let result: AccountPersonnel | undefined
    const ok = await action.run(async () => {
      result = await accountPersonnelApi.link(token, username, {
        employee_id: selected?.employee_id ?? null,
        expected_employee_id: initial.employee?.employee_id ?? null,
        expected_version: initial.employee?.version ?? null,
        target_version: selected?.version ?? null,
      })
    })
    if (ok && result) await onSaved(result)
  }
  return <Dialog open fullWidth maxWidth="sm" onClose={() => { if (!action.busy) close(onClose) }} aria-labelledby="account-register-link-title">
    <Box component="form" onSubmit={event => { event.preventDefault(); void save() }}>
      <DialogTitle id="account-register-link-title">Associer une fiche du registre</DialogTitle>
      <DialogContent dividers>
        <Alert severity="info" sx={{ mb: 2 }}>Choisissez une personne déjà enregistrée. Une fiche ne peut être associée qu’à un seul compte. Les informations et l’historique du registre sont conservés.</Alert>
        <Typography variant="body2" sx={{ mb: 1, overflowWrap: 'anywhere' }}>Compte concerné : {initial.account_display_name} · {initial.account_email}</Typography>
        <Typography variant="body2" sx={{ mb: 2 }}>Association actuelle : {initial.employee ? `${initial.employee.full_name} · n° ${initial.employee.registration_number}` : 'Aucune'}</Typography>
        <TextField fullWidth size="small" label="Rechercher par nom ou numéro" value={search} disabled={action.busy} onChange={event => setSearch(event.target.value)} sx={{ mb: 2 }} />
        {choices.error && <Alert severity="error" sx={{ mb: 2 }}>{choices.error.message}<Button onClick={() => void choices.refetch()}>Réessayer</Button></Alert>}
        <TextField fullWidth select label="Fiche du registre" value={selected?.employee_id ?? ''} disabled={action.busy || choices.isPending} onChange={event => setSelected(options.find(item => item.employee_id === Number(event.target.value)) || null)}>
          <MenuItem value="">Aucune association</MenuItem>
          {options.map(item => <MenuItem key={item.employee_id} value={item.employee_id} sx={{ whiteSpace: 'normal', overflowWrap: 'anywhere' }}>{item.full_name} · n° {item.employee_id}{item.job_title ? ` · ${item.job_title}` : ''}</MenuItem>)}
        </TextField>
        {choices.isFetching && <CircularProgress size={20} sx={{ mt: 1 }} />}
        {choices.data?.has_more && <Typography variant="caption" sx={{ display: 'block', mt: 1 }}>Affinez la recherche pour voir les autres fiches.</Typography>}
        {!selected && initial.employee && <Alert severity="warning" sx={{ mt: 2 }}>La fiche sera dissociée de ce compte. Ses données, périodes d’emploi et historiques ne seront pas supprimés.</Alert>}
        <Box sx={{ mt: 2 }}><WriteError error={action.error} conflict={action.conflict} /></Box>
      </DialogContent>
      <DialogActions><Button disabled={action.busy} onClick={() => close(onClose)}>Annuler</Button><Button variant="contained" type="submit" disabled={!changed || action.busy || action.conflict || !!choices.error}>{action.busy ? 'Enregistrement…' : 'Enregistrer l’association'}</Button></DialogActions>
    </Box>
  </Dialog>
}
