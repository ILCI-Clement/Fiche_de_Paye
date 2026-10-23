import { useState } from 'react'
import type { ReactNode } from 'react'
import { Box, Button, IconButton, Tab, Table, TableBody, TableCell, TableContainer, TableHead, TableRow, Tabs, Tooltip, Typography, useMediaQuery } from '@mui/material'
import { useTheme } from '@mui/material/styles'
import MoreHorizRounded from '@mui/icons-material/MoreHorizRounded'
import { contractLabel, dateLabel, statusLabels, typeLabel } from './api'
import type { Employee } from './api'
import { SensitiveNumber } from './SensitiveNumber'

type Sheet = 'identity' | 'employment' | 'special'
type Size = 'xs' | 'sm' | 'md' | 'lg' | 'xl'
interface Column {
  id: string; label: string; group: string; width: number | string; min?: Size;
  render: (employee: Employee) => ReactNode;
}

function CellText({ value, title }: { value?: string | null; title?: string }) {
  return <Tooltip title={title || value || ''}><Typography variant="body2" noWrap sx={{ fontSize: 'inherit' }}>{value || '—'}</Typography></Tooltip>
}

/** Template columns B:R, split into readable sheets while retaining one row per employee. */
export function PersonnelTable({ employees, token, active, resetKey, onOpen }: {
  employees: Employee[]; token: string; active: boolean; resetKey: string; onOpen: (id: number) => void;
}) {
  const [sheet, setSheet] = useState<Sheet>('identity')
  const theme = useTheme()
  const sm = useMediaQuery(theme.breakpoints.up('sm'))
  const md = useMediaQuery(theme.breakpoints.up('md'))
  const lg = useMediaQuery(theme.breakpoints.up('lg'))
  const xl = useMediaQuery(theme.breakpoints.up('xl'))
  const sizes = { xs: true, sm, md, lg, xl }
  const columns: Record<Sheet, Column[]> = {
    identity: [
      { id: 'nationality', label: 'Nationalité', group: 'Informations personnelles', width: xl ? '9%' : '8%', min: 'md', render: e => <CellText value={e.nationality} /> },
      { id: 'sex', label: 'Sexe', group: 'Informations personnelles', width: 52, min: 'lg', render: e => <CellText value={e.sex === 'F' ? 'F' : e.sex === 'M' ? 'H' : ''} title={e.sex === 'F' ? 'Femme' : e.sex === 'M' ? 'Homme' : ''} /> },
      { id: 'address', label: 'Adresse', group: 'Informations personnelles', width: xl ? '24%' : '17%', min: 'lg', render: e => <CellText value={e.address} /> },
      { id: 'nir', label: 'N° Sécurité sociale', group: 'Informations personnelles', width: sm ? 184 : 164, render: e => <SensitiveNumber key={JSON.stringify([resetKey, e.version, sheet])} inline token={token} employeeId={e.employee_id} hasNumber={e.has_social_security_number} active={active} /> },
      { id: 'birth', label: 'Date de naissance', group: 'Informations personnelles', width: 108, min: 'lg', render: e => <CellText value={dateLabel(e.birth_date)} /> },
      { id: 'job', label: 'Emploi', group: 'Situation professionnelle', width: xl ? '12%' : '10%', min: 'md', render: e => <CellText value={e.job_title} /> },
      { id: 'qualification', label: 'Qualification', group: 'Situation professionnelle', width: xl ? '12%' : '10%', min: 'lg', render: e => <CellText value={e.qualification} /> },
    ],
    employment: [
      { id: 'job', label: 'Emploi', group: 'Situation professionnelle', width: '20%', min: 'sm', render: e => <CellText value={e.job_title} /> },
      { id: 'qualification', label: 'Qualification', group: 'Situation professionnelle', width: '18%', min: 'lg', render: e => <CellText value={e.qualification} /> },
      { id: 'entry', label: 'Entrée', group: 'Dates', width: 104, render: e => <CellText value={dateLabel(e.latest_period?.entry_date)} /> },
      { id: 'exit', label: 'Sortie effective', group: 'Dates', width: 104, min: 'md', render: e => <CellText value={dateLabel(e.latest_period?.actual_exit_date)} /> },
      { id: 'contract', label: 'Type de contrat', group: 'Contrat', width: 116, min: 'xl', render: e => <CellText value={contractLabel(e.latest_period)} /> },
    ],
    special: [
      { id: 'permit_type', label: 'Type du document', group: 'Travailleurs étrangers', width: '14%', min: 'md', render: e => <CellText value={e.latest_period?.is_foreign_worker ? e.latest_period.work_permit_type : ''} /> },
      { id: 'permit_number', label: 'N° du document', group: 'Travailleurs étrangers', width: '14%', min: 'lg', render: e => <CellText value={e.latest_period?.is_foreign_worker ? e.latest_period.work_permit_number : ''} /> },
      { id: 'temporary_contract', label: 'Type de contrat', group: 'Travailleurs temporaires', width: '12%', min: 'md', render: e => <CellText value={contractLabel(e.latest_period)} /> },
      { id: 'agency', label: 'Entreprise et adresse', group: 'Travailleurs temporaires', width: '22%', min: 'xl', render: e => <CellText value={e.latest_period?.is_temporary_worker ? [e.latest_period.temporary_company_name, e.latest_period.temporary_company_address].filter(Boolean).join(' · ') : ''} /> },
      { id: 'apprentice', label: 'Apprenti', group: 'Jeunes travailleurs', width: 62, render: e => e.latest_period?.is_young_worker && e.latest_period.youth_contract_type === 'apprenti' ? 'Oui' : '—' },
      { id: 'youth_qualification', label: 'Contrat de qualification', group: 'Jeunes travailleurs', width: 82, min: 'lg', render: e => e.latest_period?.is_young_worker && e.latest_period.youth_contract_type === 'qualification' ? 'Oui' : '—' },
      { id: 'adaptation', label: 'Contrat d’adaptation', group: 'Jeunes travailleurs', width: 82, min: 'xl', render: e => e.latest_period?.is_young_worker && e.latest_period.youth_contract_type === 'adaptation' ? 'Oui' : '—' },
    ],
  }
  const visible = columns[sheet].filter(column => sizes[column.min || 'xs'])
  const groups: { label: string; count: number }[] = []
  for (const column of visible) {
    const last = groups.at(-1)
    if (last?.label === column.group) last.count += 1
    else groups.push({ label: column.group, count: 1 })
  }
  return <>
    <Tabs value={sheet} onChange={(_, value: Sheet) => setSheet(value)} variant="scrollable" scrollButtons="auto" aria-label="Feuilles du registre" sx={{ flexShrink: 0, borderBottom: 1, borderColor: 'divider', '&.MuiTabs-root': { minHeight: 38 }, '& .MuiTab-root': { minHeight: 38, py: .75, px: 2, fontSize: 12 } }}>
      <Tab id="registry-sheet-identity" aria-controls="registry-table" value="identity" label="Informations de base" />
      <Tab id="registry-sheet-employment" aria-controls="registry-table" value="employment" label="Emploi et dates" title="Dernière période d’emploi. L’historique complet reste disponible dans la fiche." />
      <Tab id="registry-sheet-special" aria-controls="registry-table" value="special" label="Situations particulières" title="Situations de la dernière période d’emploi. L’historique complet reste disponible dans la fiche." />
    </Tabs>
    <TableContainer id="registry-table" role="tabpanel" aria-labelledby={`registry-sheet-${sheet}`} sx={{ flex: 1, minHeight: 0, overflowY: 'auto', overflowX: 'hidden', scrollbarGutter: 'stable' }}>
      <Table stickyHeader size="small" aria-label="Registre du personnel" sx={{ width: '100%', tableLayout: 'fixed', borderCollapse: 'separate', borderSpacing: 0,
        '& .MuiTableCell-root': { overflow: 'hidden', borderRight: 1, borderRightColor: 'divider', lineHeight: 1.4 },
        '& .MuiTableCell-root:last-child': { borderRight: 0 },
        '& .MuiTableCell-head': { bgcolor: 'background.paper', whiteSpace: 'normal', overflowWrap: 'anywhere' },
        '& .registry-group .MuiTableCell-root': { top: 0, height: 36, py: .5, fontSize: 12 },
        '& .registry-columns .MuiTableCell-root': { top: 36, height: 46, py: .75, fontSize: 12 },
        '& .MuiTableBody-root .MuiTableRow-root:nth-of-type(even)': { bgcolor: 'action.hover' },
      }}>
        <colgroup><col />{visible.map(column => <col key={column.id} style={{ width: column.width }} />)}<col style={{ width: sm ? 44 : 36 }} /></colgroup>
        <TableHead>
          <TableRow className="registry-group">
            <TableCell rowSpan={2} scope="col">Nom et prénoms</TableCell>
            {groups.map(group => <TableCell key={group.label} colSpan={group.count} scope="colgroup" align="center"><Tooltip title={group.label}><Box component="span" sx={{ display: 'block', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', fontWeight: 600 }}>{group.label}</Box></Tooltip></TableCell>)}
            <TableCell rowSpan={2} scope="col" aria-label="Fiche complète" />
          </TableRow>
          <TableRow className="registry-columns">{visible.map(column => <TableCell key={column.id} scope="col" align={sheet === 'special' && column.group === 'Jeunes travailleurs' ? 'center' : 'left'}>{column.label}</TableCell>)}</TableRow>
        </TableHead>
        <TableBody>{employees.map(employee => <TableRow key={employee.employee_id} hover>
          <TableCell component="th" scope="row">
            <Tooltip title={employee.full_name}><Button onClick={() => onOpen(employee.employee_id)} sx={{ minWidth: 0, maxWidth: '100%', minHeight: 28, p: 0, textAlign: 'left', justifyContent: 'flex-start' }}><Box component="span" sx={{ minWidth: 0, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{employee.full_name}</Box></Button></Tooltip>
            {sheet === 'employment' && <Typography variant="caption" color="text.secondary" sx={{ display: 'block', fontSize: 11 }}>{typeLabel(employee.personnel_type)} · {statusLabels[employee.status] || 'À préciser'}</Typography>}
            {employee.is_archived && <Typography variant="caption" color="text.secondary" sx={{ display: 'block' }}>Archivé</Typography>}
            {!!employee.missing_fields?.length && <Tooltip title={`À compléter : ${employee.missing_fields.join(', ')}`}><Typography variant="caption" color="text.secondary" sx={{ display: 'block', fontSize: 11 }}>À compléter</Typography></Tooltip>}
          </TableCell>
          {visible.map(column => <TableCell key={column.id} align={sheet === 'special' && column.group === 'Jeunes travailleurs' ? 'center' : 'left'}>{column.render(employee)}</TableCell>)}
          <TableCell align="center" sx={{ px: .25 }}><Tooltip title="Ouvrir la fiche complète"><IconButton size="small" aria-label={`Ouvrir la fiche complète de ${employee.full_name}`} onClick={() => onOpen(employee.employee_id)}><MoreHorizRounded fontSize="small" /></IconButton></Tooltip></TableCell>
        </TableRow>)}</TableBody>
      </Table>
    </TableContainer>
  </>
}
