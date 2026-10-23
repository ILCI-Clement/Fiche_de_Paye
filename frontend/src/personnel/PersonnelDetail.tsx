import { personDisplayName } from '../personDisplayName'
import { useId, useState } from 'react'
import { Alert, Box, Button, Chip, CircularProgress, Dialog, DialogActions, DialogContent, DialogTitle, IconButton, Tab, Table, TableBody, TableCell, TableContainer, TableHead, TableRow, Tabs, Tooltip, Typography } from '@mui/material'
import EditOutlined from '@mui/icons-material/EditOutlined'
import AddRounded from '@mui/icons-material/AddRounded'
import ArchiveOutlined from '@mui/icons-material/ArchiveOutlined'
import UnarchiveOutlined from '@mui/icons-material/UnarchiveOutlined'
import CloseRounded from '@mui/icons-material/CloseRounded'
import { useQuery, useQueryClient } from '@tanstack/react-query'
import { DeleteAction } from '../components/DeleteAction'
import { PermanentDeleteDialog } from '../components/PermanentDeleteDialog'
import { contractLabel, dateLabel, personnelApi, statusLabels, typeLabel } from './api'
import type { Employee, EmploymentPeriod } from './api'
import { PersonnelEditor } from './PersonnelEditor'
import { EmploymentEditor } from './EmploymentEditor'
import { PersonnelHistory } from './PersonnelHistory'
import { ArchiveConfirmation } from './ArchiveConfirmation'
import { useRegistryViewer } from './context'
import { SensitiveNumber } from './SensitiveNumber'
import { EmploymentSummary } from './EmploymentSummary'

export function PersonnelDetail({ token, employeeId, onClose, onSaved, canDelete = false }: { token: string; employeeId: number; onClose: () => void; onSaved: (message: string) => void; canDelete?: boolean }) {
  const viewer = useRegistryViewer()
  const client = useQueryClient()
  const detail = useQuery({ queryKey: ['personnel-register', viewer, 'detail', employeeId], queryFn: ({ signal }) => personnelApi.detail(token, employeeId, signal), retry: false })
  const [tab, setTab] = useState<'information' | 'periods' | 'history'>('information')
  const [editing, setEditing] = useState<Employee | null>(null)
  const [employment, setEmployment] = useState<{ employee: Employee; period?: EmploymentPeriod; situationsOnly?: boolean } | null>(null)
  const [archiving, setArchiving] = useState<Employee | null>(null)
  const [deleting, setDeleting] = useState<Employee | null>(null)
  const tabsId = useId()
  const employee = detail.data?.employee
  const latestPeriod = detail.data?.periods[0]
  const inEditor = !!editing || !!employment || !!archiving || !!deleting
  const rows: [string, string | null | undefined][] = employee ? [
    ['Nom et prénoms', employee.full_name], ['Nom', employee.last_name], ['Prénoms', employee.first_name], ['Type de personnel', typeLabel(employee.personnel_type)],
    ['Nationalité', employee.nationality], ['Sexe', employee.sex === 'F' ? 'Femme' : employee.sex === 'M' ? 'Homme' : ''],
    ['Date de naissance', dateLabel(employee.birth_date)], ['Adresse', employee.address], ['Emploi', employee.job_title],
    ['Qualification', employee.qualification], ['Compte associé', employee.linked_username ? personDisplayName({ display_name: employee.linked_display_name }) : 'Aucun compte associé'],
  ] : []
  return <>
    <Dialog open={!inEditor} fullScreen onClose={onClose} aria-labelledby="personnel-detail-title"
      slotProps={{ paper: { sx: {
        m: 0, height: '100dvh', maxHeight: '100dvh', width: '100%', borderRadius: 0,
      } } }}>
      <DialogTitle id="personnel-detail-title" sx={{ pb: 1, flexShrink: 0 }}>
        <Box sx={{ display: 'flex', alignItems: 'center', gap: 1, flexWrap: 'wrap' }}>
          <Typography component="span" variant="h6" sx={{ flex: 1, minWidth: 0, overflowWrap: 'anywhere' }}>{employee?.full_name || 'Fiche du personnel'}</Typography>
          {employee && <><Chip size="small" variant="outlined" label={`N° ${employee.registration_number}`} />{employee.is_archived && <Chip size="small" label="Archivé" />}</>}
          <Tooltip title="Fermer la fiche"><IconButton aria-label="Fermer la fiche" onClick={onClose}><CloseRounded /></IconButton></Tooltip>
        </Box>
      </DialogTitle>
      <DialogContent dividers sx={{ flex: 1, minHeight: 0, minWidth: 0, overflowY: 'auto', scrollbarGutter: 'stable' }}>
        {detail.isPending ? <CircularProgress size={28} /> : detail.error ? <Alert severity="error">{detail.error.message}<Button onClick={() => void detail.refetch()}>Réessayer</Button></Alert> : employee && <>
          <Box sx={{ display: 'flex', alignItems: 'center', flexWrap: 'wrap', gap: 1, mb: 2 }}>
            <Chip size="small" color={employee.status === 'employed' ? 'primary' : 'default'} variant="outlined" label={statusLabels[employee.status] || 'À préciser'} />
            <Typography variant="caption" color="text.secondary" sx={{ flex: 1 }}>État issu de la dernière période datée · le compte de connexion est géré séparément.</Typography>
            <Button disabled={detail.isFetching} onClick={() => void detail.refetch()}>Actualiser</Button>
          </Box>
          <Tabs value={tab} onChange={(_, next: typeof tab) => setTab(next)} variant="scrollable" scrollButtons="auto" aria-label="Sections de la fiche du personnel" sx={{ mb: 2, borderBottom: 1, borderColor: 'divider' }}>
            <Tab value="information" label="Informations" id={`${tabsId}-information-tab`} aria-controls={`${tabsId}-information-panel`} />
            <Tab value="periods" label={`Périodes d’emploi (${detail.data.periods.length})`} id={`${tabsId}-periods-tab`} aria-controls={`${tabsId}-periods-panel`} />
            <Tab value="history" label="Historique des modifications" id={`${tabsId}-history-tab`} aria-controls={`${tabsId}-history-panel`} />
          </Tabs>
          {!!employee.missing_fields?.length && <Alert severity="info" sx={{ mb: 2 }}>Fiche à compléter : {employee.missing_fields.join(', ')}. Les autres opérations restent disponibles.</Alert>}
          {tab === 'information' && <Box role="tabpanel" id={`${tabsId}-information-panel`} aria-labelledby={`${tabsId}-information-tab`}>
            <Box sx={{ display: 'flex', justifyContent: 'flex-end', mb: 1 }}><Button startIcon={<EditOutlined />} onClick={() => setEditing({ ...employee })}>Modifier les informations</Button></Box>
            <Box sx={{ border: 1, borderColor: 'divider', borderRadius: 1.5, p: 2, mb: 3 }}>
              <Box sx={{ display: 'flex', gap: 1, alignItems: 'center', flexWrap: 'wrap', mb: 1 }}>
                <Box sx={{ flex: 1, minWidth: 180 }}>
                  <Typography variant="subtitle2">Situations particulières</Typography>
                  <Typography variant="caption" color="text.secondary">{latestPeriod ? `Dernière période · entrée le ${dateLabel(latestPeriod.entry_date)}` : 'Aucune période d’emploi renseignée'}</Typography>
                </Box>
                {detail.data.periods.length > 1 && <Button size="small" onClick={() => setTab('periods')}>Toutes les périodes</Button>}
                <Button size="small" startIcon={<EditOutlined />} onClick={() => setEmployment({ employee: { ...employee }, period: latestPeriod ? { ...latestPeriod } : undefined, situationsOnly: !!latestPeriod })}>{latestPeriod ? 'Modifier les situations particulières' : 'Renseigner la situation'}</Button>
              </Box>
              {latestPeriod && <>
                <Box sx={{ display: 'flex', gap: 1, flexWrap: 'wrap' }}>
                  {[
                    { label: 'Travailleur étranger', selected: latestPeriod.is_foreign_worker },
                    { label: 'Travailleur temporaire', selected: latestPeriod.is_temporary_worker },
                    { label: 'Jeune travailleur / Apprenti', selected: latestPeriod.is_young_worker },
                  ].map(item => <Chip key={item.label} size="small" variant="outlined" color={item.selected ? 'primary' : 'default'} label={`${item.label} : ${item.selected ? 'Oui' : 'Non'}`} />)}
                </Box>
                <EmploymentSummary period={latestPeriod} />
              </>}
            </Box>
            <Box component="dl" sx={{ m: 0, display: 'grid', gridTemplateColumns: { xs: '1fr', sm: '1fr 1fr' }, gap: 2.5 }}>
              {rows.map(([label, value]) => <Box component="div" key={label} sx={label === 'Adresse' ? { gridColumn: '1 / -1' } : undefined}><Typography component="dt" variant="caption" color="text.secondary">{label}</Typography><Typography component="dd" variant="body2" sx={{ m: 0, mt: .5, whiteSpace: 'pre-wrap', overflowWrap: 'anywhere' }}>{value || 'Non renseigné'}</Typography></Box>)}
              <Box component="div"><Typography component="dt" variant="caption" color="text.secondary">N° de sécurité sociale</Typography><Box component="dd" sx={{ m: 0, mt: .5 }}>{!inEditor && <SensitiveNumber key={`${employee.employee_id}-${employee.version}`} token={token} employeeId={employee.employee_id} hasNumber={!!employee.has_social_security_number} />}</Box></Box>
            </Box>
          </Box>}
          {tab === 'periods' && <Box role="tabpanel" id={`${tabsId}-periods-panel`} aria-labelledby={`${tabsId}-periods-tab`}>
            <Box sx={{ display: 'flex', gap: 2, alignItems: 'center', flexWrap: 'wrap', mb: 2 }}><Typography variant="body2" color="text.secondary" sx={{ flex: 1 }}>Les périodes précédentes sont conservées lors d’une nouvelle entrée. Une correction reste visible dans l’historique.</Typography><Button variant="outlined" startIcon={<AddRounded />} onClick={() => setEmployment({ employee: { ...employee } })}>Ajouter une période</Button></Box>
            {!detail.data.periods.length ? <Alert severity="info">Aucune période renseignée. Complétez la date d’entrée depuis « Ajouter une période ».</Alert> : <TableContainer sx={{ overflowX: 'hidden' }}><Table size="small" sx={{ tableLayout: 'fixed', '& td, & th': { overflowWrap: 'anywhere', px: 1 } }} aria-label="Toutes les périodes d’emploi"><TableHead><TableRow><TableCell>Période et situations</TableCell><TableCell>Date d’entrée</TableCell><TableCell>Sortie effective</TableCell><TableCell>Fin prévue du contrat</TableCell><TableCell>Contrat</TableCell><TableCell align="right" sx={{ width: 64 }}>Action</TableCell></TableRow></TableHead><TableBody>{detail.data.periods.map(period => <TableRow key={period.period_id}>
              <TableCell>{period.account_contract_id ? `Contrat #${period.account_contract_id}` : `#${period.period_id}`}<EmploymentSummary period={period} /></TableCell>
              <TableCell>{dateLabel(period.entry_date)}</TableCell><TableCell>{dateLabel(period.actual_exit_date)}</TableCell><TableCell>{dateLabel(period.exit_date)}</TableCell><TableCell>{contractLabel(period)}</TableCell>
              <TableCell align="right"><Tooltip title="Corriger cette période"><IconButton aria-label={`Corriger la période ${period.period_id}`} onClick={() => setEmployment({ employee: { ...employee }, period: { ...period } })}><EditOutlined fontSize="small" /></IconButton></Tooltip></TableCell>
            </TableRow>)}</TableBody></Table></TableContainer>}
          </Box>}
          {tab === 'history' && <Box role="tabpanel" id={`${tabsId}-history-panel`} aria-labelledby={`${tabsId}-history-tab`}><PersonnelHistory token={token} employeeId={employeeId} /></Box>}
        </>}
      </DialogContent>
      <DialogActions sx={{ px: 3, py: 2, flexShrink: 0 }}>
        {employee && <Button startIcon={employee.is_archived ? <UnarchiveOutlined /> : <ArchiveOutlined />} onClick={() => setArchiving({ ...employee })} sx={{ mr: 'auto' }}>{employee.is_archived ? 'Restaurer la fiche' : 'Archiver la fiche'}</Button>}
        {employee?.is_archived && canDelete && <DeleteAction label="Supprimer définitivement cette fiche" onClick={() => setDeleting({ ...employee })} />}
        <Button onClick={onClose}>Fermer</Button>
      </DialogActions>
    </Dialog>
    {editing && <PersonnelEditor token={token} employee={editing} onClose={() => setEditing(null)} onSaved={onSaved} />}
    {employment && <EmploymentEditor token={token} employee={employment.employee} period={employment.period} situationsOnly={employment.situationsOnly} onClose={() => setEmployment(null)} onSaved={onSaved} />}
    {archiving && <ArchiveConfirmation token={token} employee={archiving} onClose={() => setArchiving(null)} onSaved={onSaved} />}
    {deleting && <PermanentDeleteDialog title="Supprimer définitivement cette fiche du personnel ?" expected={String(deleting.employee_id)} description={`La fiche de ${deleting.full_name}, ses périodes d’emploi et son historique du registre seront supprimés. Le compte de connexion est géré séparément.`}
      onConfirm={confirmation => personnelApi.remove(token, deleting.employee_id, deleting.version, confirmation)} onClose={() => setDeleting(null)} onDeleted={() => {
        onClose(); onSaved('Fiche du personnel supprimée.')
        client.removeQueries({ queryKey: ['personnel-register', viewer, 'detail', employeeId] })
        client.removeQueries({ queryKey: ['personnel-register', viewer, 'history', employeeId] })
        void client.invalidateQueries({ queryKey: ['personnel-register', viewer, 'list'] })
      }} />}
  </>
}
