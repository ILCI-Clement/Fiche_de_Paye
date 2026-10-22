import { Button, IconButton, Tooltip } from '@mui/material'
import UploadFileRounded from '@mui/icons-material/UploadFileRounded'
import RefreshRounded from '@mui/icons-material/RefreshRounded'
import { WorkspaceToolbar } from '../components/WorkspaceToolbar'
import { PeriodToolbar } from '../attendance/PeriodToolbar'

export function DocumentsToolbar({ month, year, busy, canUpload, refreshing, onPeriod, onUpload, onRefresh }: {
  month: number; year: number; busy: boolean; canUpload: boolean; refreshing: boolean;
  onPeriod: (month: number, year: number) => void; onUpload: () => void; onRefresh: () => void;
}) {
  return <WorkspaceToolbar>
    <PeriodToolbar label="Mois de dépôt et de consultation" month={month} year={year} disabled={busy} onChange={onPeriod} />
    <Button variant="contained" startIcon={<UploadFileRounded />} disabled={busy || !canUpload} onClick={onUpload} sx={{ whiteSpace: 'nowrap', flexShrink: 0 }}>Déposer un document</Button>
    <Tooltip title="Actualiser les documents"><span><IconButton aria-label="Actualiser les documents" disabled={busy || refreshing} onClick={onRefresh}><RefreshRounded fontSize="small" /></IconButton></span></Tooltip>
  </WorkspaceToolbar>
}
