import { useState } from 'react'
import { Alert, CircularProgress, IconButton, Tooltip } from '@mui/material'
import DownloadOutlined from '@mui/icons-material/DownloadOutlined'
import { downloadFile } from '../download'
import { medicalApi } from './api'
import type { MedicalPeriod } from './api'

export function MedicalExport({ token, period, q = '', archived, employeeId, ids, disabled = false, allHistory = false, name }: {
  token: string; period: MedicalPeriod; q?: string; archived: boolean; employeeId?: number; ids?: number[]; disabled?: boolean; allHistory?: boolean; name?: string;
}) {
  const [busy, setBusy] = useState(false), [error, setError] = useState('')
  const label = allHistory ? `Exporter toutes les attestations de ${name || 'ce salarié'}` : ids ? `Exporter la sélection (${ids.length})` : 'Exporter les attestations filtrées — toutes les pages'
  async function run() {
    setBusy(true); setError('')
    try { downloadFile(await medicalApi.export(token, period, q, archived, employeeId, 'zip', ids, allHistory)) }
    catch (reason) { setError(reason instanceof Error ? reason.message : 'Export impossible.') }
    finally { setBusy(false) }
  }
  return <>
    <Tooltip title={`${label} (ZIP)`}><span><IconButton aria-label={label} disabled={busy || disabled} onClick={() => void run()}>{busy ? <CircularProgress size={20} /> : <DownloadOutlined />}</IconButton></span></Tooltip>
    {error && <Alert severity="error" onClose={() => setError('')}>{error}</Alert>}
  </>
}
