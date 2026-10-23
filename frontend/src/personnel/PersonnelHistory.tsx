import { useState } from 'react'
import { Accordion, AccordionDetails, AccordionSummary, Alert, Box, Button, CircularProgress, Table, TableBody, TableCell, TableContainer, TableHead, TablePagination, TableRow, Typography } from '@mui/material'
import ExpandMoreRounded from '@mui/icons-material/ExpandMoreRounded'
import { useQuery } from '@tanstack/react-query'
import { dateLabel, personnelApi, typeLabel, contractLabel, youthContractLabel } from './api'
import type { PersonnelType, RegistryEvent, RegistrySnapshot } from './api'
import { useRegistryViewer } from './context'
import { personDisplayName } from '../personDisplayName'

const labels: Record<string, string> = { full_name: 'Nom et prénoms', last_name: 'Nom', first_name: 'Prénoms', nationality: 'Nationalité', sex: 'Sexe', address: 'Adresse', birth_date: 'Date de naissance', job_title: 'Emploi', qualification: 'Qualification', personnel_type: 'Type de personnel', linked_username: 'Compte associé', is_archived: 'Archivage', has_social_security_number: 'N° de sécurité sociale · présence' }
const actions: Record<string, string> = { contract_dates_synced: 'Synchronisation des dates du contrat', excel_exported: 'Export Excel du registre', registered: 'Inscription au registre', details_updated: 'Correction des informations', account_renamed: 'Renommage du compte associé', account_linked: 'Association à un compte', account_unlinked: 'Dissociation du compte', archived: 'Archivage', restored: 'Restauration', period_added: 'Ajout d’une période d’emploi', period_updated: 'Correction d’une période d’emploi', social_security_number_changed: 'Modification du numéro de sécurité sociale', social_security_number_viewed: 'Consultation du numéro de sécurité sociale' }
function snapshot(value: RegistrySnapshot | null) {
  const fields = new Map<string, { label: string; value: string; comparison?: string }>()
  if (!value) return fields
  for (const [key, label] of Object.entries(labels)) {
    const raw = value.employee?.[key as keyof RegistrySnapshot['employee']]
    if (raw === undefined) continue
    const text = key === 'linked_username' ? raw ? personDisplayName({ display_name: value.employee?.linked_display_name }) : '' : key === 'personnel_type' ? typeLabel(raw as PersonnelType) : key === 'birth_date' ? dateLabel(raw as string | null) : key === 'sex' ? raw === 'F' ? 'Femme' : raw === 'M' ? 'Homme' : '' : key === 'is_archived' ? raw ? 'Archivé' : 'Non archivé' : key === 'has_social_security_number' ? raw ? 'Renseigné' : 'Non renseigné' : String(raw ?? '')
    fields.set(`employee.${key}`, { label, value: text, comparison: key === 'linked_username' ? String(raw ?? '') : text })
  }
  for (const period of value.periods || []) {
    const title = `Période #${period.period_id}`
    fields.set(`period.${period.period_id}.entry`, { label: `${title} · Date d’entrée`, value: dateLabel(period.entry_date) })
    fields.set(`period.${period.period_id}.exit`, { label: `${title} · Fin prévue du contrat`, value: dateLabel(period.exit_date) })
    fields.set(`period.${period.period_id}.actualExit`, { label: `${title} · Date de sortie effective`, value: dateLabel(period.actual_exit_date) })
    fields.set(`period.${period.period_id}.contract`, { label: `${title} · Contrat`, value: contractLabel(period) })
    for (const [key, label] of Object.entries({ is_foreign_worker: 'Travailleur étranger', work_permit_type: 'Type du document', work_permit_number: 'N° du document',
      is_temporary_worker: 'Travailleur temporaire', temporary_company_name: 'Entreprise de travail temporaire', temporary_company_address: 'Adresse de l’entreprise',
      is_young_worker: 'Jeune travailleur / Apprenti', youth_contract_type: 'Contrat spécifique' })) {
      const raw = period[key as keyof typeof period]
      if (raw === undefined) continue
      const text = key === 'youth_contract_type' ? youthContractLabel(period.youth_contract_type) : key.startsWith('is_') ? raw ? 'Oui' : 'Non' : String(raw ?? '')
      fields.set(`period.${period.period_id}.${key}`, { label: `${title} · ${label}`, value: text })
    }
  }
  return fields
}
function changes(event: RegistryEvent) {
  const before = snapshot(event.before), after = snapshot(event.after)
  return [...new Set([...before.keys(), ...after.keys()])].flatMap(key => {
    const old = before.get(key), next = after.get(key)
    if ((old?.comparison ?? old?.value) === (next?.comparison ?? next?.value)) return []
    // An initial empty field adds no information to the creation record.
    if (!event.before && (!next?.value || ['—', 'Non', 'Non renseigné', 'Non archivé'].includes(next.value))) return []
    return [{ key, label: next?.label || old!.label, before: old?.value || 'Non renseigné', after: next?.value || 'Non renseigné' }]
  })
}
const timestamp = (value: string) => {
  const normalized = /(?:Z|[+-]\d\d:\d\d)$/.test(value) ? value : `${value.replace(' ', 'T')}Z`
  const date = new Date(normalized)
  return Number.isNaN(date.getTime()) ? value : date.toLocaleString('fr-FR', { timeZone: 'Europe/Paris', dateStyle: 'short', timeStyle: 'short' })
}

export function PersonnelHistory({ token, employeeId }: { token: string; employeeId: number }) {
  const [page, setPage] = useState(0)
  const viewer = useRegistryViewer()
  const history = useQuery({ queryKey: ['personnel-register', viewer, 'history', employeeId, page], queryFn: ({ signal }) => personnelApi.history(token, employeeId, page + 1, signal), retry: false })
  return <Box>
    <Typography variant="body2" color="text.secondary" sx={{ mb: 2 }}>Les inscriptions, corrections, associations de comptes et changements d’archivage sont conservés avec leur auteur. Les anciennes versions ne sont pas remplacées.</Typography>
    {history.isPending ? <CircularProgress size={24} /> : history.error ? <Alert severity="error">{history.error.message}<Button onClick={() => void history.refetch()}>Réessayer</Button></Alert> : <>
      {!history.data.events.length && <Typography color="text.secondary">Aucune modification enregistrée.</Typography>}
      {history.data.events.map(event => {
        const rows = changes(event)
        const sensitive = event.action === 'excel_exported' || event.action === 'social_security_number_changed' || event.action === 'social_security_number_viewed'
        return <Accordion key={event.event_id} disableGutters sx={{ mb: 1 }}>
          <AccordionSummary expandIcon={<ExpandMoreRounded />}>
            <Box><Typography sx={{ fontWeight: 500 }}>{actions[event.action] || 'Modification du registre'}</Typography><Typography variant="caption" color="text.secondary">{timestamp(event.created_at)} · {personDisplayName({ display_name: event.actor_display_name })} · {sensitive ? 'Opération sensible' : `${rows.length} champ(s)`}</Typography></Box>
          </AccordionSummary>
          <AccordionDetails>
            {sensitive && <Typography variant="body2" color="text.secondary" sx={{ mb: rows.length ? 1 : 0 }}>{event.action === 'excel_exported' ? `Export de ${event.after.export?.row_count ?? 1} ligne(s) au format Excel, avec le numéro de sécurité sociale lorsqu’il est renseigné.` : event.action === 'social_security_number_viewed' ? 'Un accès au numéro a été autorisé et enregistré.' : 'Le numéro a été ajouté, remplacé ou effacé. Aucun ancien ou nouveau numéro n’est conservé dans cet historique.'} Les valeurs ne sont jamais affichées ici.</Typography>}
            {!rows.length ? !sensitive && <Typography variant="body2" color="text.secondary">État enregistré sans changement des champs affichés.</Typography> : <TableContainer><Table size="small" aria-label="Anciennes et nouvelles valeurs"><TableHead><TableRow><TableCell>Champ</TableCell><TableCell>Avant</TableCell><TableCell>Après</TableCell></TableRow></TableHead><TableBody>{rows.map(row => <TableRow key={row.key}><TableCell sx={{ minWidth: 160 }}>{row.label}</TableCell><TableCell sx={{ minWidth: 150, maxWidth: 360, overflowWrap: 'anywhere', whiteSpace: 'pre-wrap', color: 'text.secondary' }}>{row.before}</TableCell><TableCell sx={{ minWidth: 150, maxWidth: 360, overflowWrap: 'anywhere', whiteSpace: 'pre-wrap' }}>{row.after}</TableCell></TableRow>)}</TableBody></Table></TableContainer>}
          </AccordionDetails>
        </Accordion>
      })}
      <TablePagination component="div" count={history.data.total} page={page} rowsPerPage={history.data.page_size} rowsPerPageOptions={[]} onPageChange={(_, next) => setPage(next)} labelDisplayedRows={({ from, to, count }) => `${from}–${to} sur ${count}`} getItemAriaLabel={type => type === 'next' ? 'Page suivante de l’historique' : type === 'previous' ? 'Page précédente de l’historique' : type === 'first' ? 'Première page de l’historique' : 'Dernière page de l’historique'} />
    </>}
  </Box>
}
