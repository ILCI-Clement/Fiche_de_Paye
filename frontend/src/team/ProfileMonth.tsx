import { Box, Chip, Divider, IconButton, Paper, Tooltip, Typography } from '@mui/material'
import VisibilityOutlined from '@mui/icons-material/VisibilityOutlined'
import { signedStatus } from '../attendance/signedFichesApi'

export interface MonthlyProfile {
  month: number; attendance_sheets: number; documents: number; archived_documents: number;
  leaves: { id: number; start_date: string; end_date: string; start_slot: string; end_slot: string; status: string }[];
  signature_requests: { document_id: string; mode: string; created_at: string; status: string; archive_id?: string }[];
}
const monthName = (month: number) => new Intl.DateTimeFormat('fr-FR', { month: 'long' }).format(new Date(2026, month - 1, 1))
const statusLabels: Record<string, string> = { pending: 'En attente', approved: 'Approuvé', rejected: 'Refusé', withdrawn: 'Retiré' }
const slotLabels: Record<string, string> = { full: 'journée', morning: 'matin', afternoon: 'après-midi' }
const dateLabel = (value: string) => new Date(`${value.slice(0, 10)}T12:00:00`).toLocaleDateString('fr-FR')

export function ProfileMonth({ item, year, current, signaturesAvailable, onPreview }: {
  item: MonthlyProfile; year: number; current: boolean; signaturesAvailable: boolean; onPreview: (id: string, month: number) => void
}) {
  const pending = item.leaves.filter(leave => leave.status === 'pending').length
  const approved = item.leaves.filter(leave => leave.status === 'approved').length
  const archived = item.signature_requests.filter(signature => signature.status === 'ready' && signature.mode === 'production').length
  const titleId = `profile-month-${year}-${item.month}`
  const metrics = [
    { label: 'Présences', value: `${item.attendance_sheets} fiche(s)` },
    { label: 'Documents', value: String(item.documents) },
    { label: 'Signatures', value: signaturesAvailable ? `${archived} / ${item.signature_requests.length} archivées` : 'Indisponible' },
    { label: 'Congés', value: `${pending} en attente · ${approved} approuvés`, pending: pending > 0 },
  ]
  return <Paper component="article" variant="outlined" aria-labelledby={titleId}
    sx={{ minWidth: 0, minHeight: 212, display: 'flex', flexDirection: 'column', overflow: 'hidden', borderColor: current ? 'primary.main' : 'divider' }}>
    <Box sx={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 1, px: 1.5, py: 1.25, bgcolor: 'action.hover' }}>
      <Typography id={titleId} component="h3" variant="subtitle1" sx={{ fontWeight: 600, textTransform: 'capitalize' }}>{monthName(item.month)}</Typography>
      {current ? <Chip size="small" color="primary" variant="outlined" label="Ce mois" /> : <Typography variant="caption" color="text.secondary">{year}</Typography>}
    </Box>
    <Box sx={{ p: 1.5, display: 'grid', gridTemplateColumns: { xs: 'repeat(2, minmax(0, 1fr))', sm: 'repeat(4, minmax(0, 1fr))' }, gap: 1.5 }}>
      {metrics.map(metric => <Box key={metric.label} sx={{ minWidth: 0 }}>
        <Typography variant="caption" color="text.secondary">{metric.label}</Typography>
        <Typography variant="body2" color={metric.pending ? 'warning.main' : 'text.primary'} sx={{ mt: .5, fontWeight: 500 }}>{metric.value}</Typography>
      </Box>)}
    </Box>
    <Divider />
    <Box sx={{ px: 1.5, py: 1.25, flex: 1, minWidth: 0, maxHeight: 200, overflowY: 'auto', bgcolor: 'action.hover' }}>
      <Typography variant="caption" color="text.secondary" sx={{ display: 'block', mb: .5 }}>
        {Math.max(0, item.documents - item.archived_documents)} document(s) à traiter · {item.archived_documents} archivé(s)
      </Typography>
      {!item.leaves.length && !item.signature_requests.length && <Typography variant="caption" color="text.secondary">{signaturesAvailable ? 'Aucun congé ni envoi de signature enregistré.' : 'Aucun congé enregistré.'}</Typography>}
      {item.leaves.map(leave => <Box key={leave.id} sx={{ py: .5 }}>
        <Box sx={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: 1 }}>
          <Typography variant="caption" sx={{ fontWeight: 500 }}>Congé</Typography>
          <Typography variant="caption" color={leave.status === 'pending' ? 'warning.main' : 'text.secondary'}>{statusLabels[leave.status] || leave.status}</Typography>
        </Box>
        <Typography variant="caption" color="text.secondary" sx={{ display: 'block', overflowWrap: 'anywhere' }}>
          {dateLabel(leave.start_date)} ({slotLabels[leave.start_slot] || leave.start_slot}) → {dateLabel(leave.end_date)} ({slotLabels[leave.end_slot] || leave.end_slot})
        </Typography>
      </Box>)}
      {item.signature_requests.map((signature, index) => <Box key={`${signature.document_id}-${index}`} sx={{ py: .5, minWidth: 0 }}>
        <Box sx={{ display: 'flex', alignItems: 'center', gap: .5 }}>
          <Typography variant="caption" sx={{ flex: 1, minWidth: 0, overflowWrap: 'anywhere' }}>
            {signature.archive_id ? signedStatus[signature.status] || signature.status : 'Envoi reçu · signature à vérifier'}
            {signature.mode === 'sandbox-only' && ' · Test'}
          </Typography>
          {signature.archive_id && signature.status === 'ready' && signature.mode === 'production' && <Tooltip title="Voir le PDF signé et télécharger">
            <IconButton size="small" aria-label={`Voir le PDF signé de ${monthName(item.month)}`} onClick={() => onPreview(signature.archive_id!, item.month)}><VisibilityOutlined fontSize="small" /></IconButton>
          </Tooltip>}
        </Box>
        <Typography variant="caption" color="text.secondary" title={signature.document_id} sx={{ display: 'block', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{signature.document_id} · {dateLabel(signature.created_at)} UTC</Typography>
      </Box>)}
    </Box>
  </Paper>
}
