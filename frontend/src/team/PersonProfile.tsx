import { useState } from 'react'
import { Alert, Box, Button, Chip, CircularProgress, Dialog, DialogActions, DialogContent, DialogTitle, MenuItem, Paper, Table, TableBody, TableCell, TableContainer, TableHead, TableRow, TextField, Typography } from '@mui/material'
import { useQuery } from '@tanstack/react-query'
import { request } from '../api'
import { DepartmentBadges } from '../components/DepartmentBadge'

interface MonthlyProfile {
  month: number; attendance_sheets: number; documents: number; archived_documents: number;
  leaves: { id: number; start_date: string; end_date: string; start_slot: string; end_slot: string; status: string }[];
  signature_requests: { document_id: string; mode: string; created_at: string; status: string }[];
}
interface Profile {
  username: string; name: string; groups: string[]; manager: string | null; year: number; months: MonthlyProfile[];
  legacy_unbound_sheets: number; attendance_incomplete: boolean; signatures_available: boolean;
}
const monthName = (month: number) => new Intl.DateTimeFormat('fr-FR', { month: 'long' }).format(new Date(2026, month - 1, 1))
const statusLabels: Record<string, string> = { pending: 'En attente', approved: 'Approuvé', rejected: 'Refusé', withdrawn: 'Retiré' }
const slotLabels: Record<string, string> = { full: 'journée', morning: 'matin', afternoon: 'après-midi' }

export function PersonProfile({ username, token, viewer, onClose }: { username: string; token: string; viewer: string; onClose: () => void }) {
  const [year, setYear] = useState(() => new Date().getFullYear())
  const [month, setMonth] = useState(() => new Date().getMonth() + 1)
  const profile = useQuery({ queryKey: ['person-profile', viewer, username, year], queryFn: ({ signal }) => request<Profile>(`/person-profiles/${encodeURIComponent(username)}?year=${year}`, token, { signal }), retry: false })
  const data = profile.data
  const detail = data?.months.find(item => item.month === month)
  return <Dialog open fullWidth maxWidth="lg" onClose={onClose} aria-labelledby="person-profile-title">
    <DialogTitle id="person-profile-title">Profil · {data?.name || username}</DialogTitle>
    <DialogContent dividers>
      <Box sx={{ display: 'flex', gap: 2, flexWrap: 'wrap', alignItems: 'center', mb: 2 }}>
        <TextField select label="Année" value={year} onChange={event => setYear(Number(event.target.value))} sx={{ width: 120 }}>
          {Array.from({ length: 101 }, (_, i) => 2000 + i).reverse().map(value => <MenuItem key={value} value={value}>{value}</MenuItem>)}
        </TextField>
        {data && <><DepartmentBadges names={data.groups} /><Typography variant="body2" color="text.secondary">Responsable direct : {data.manager || 'Non renseigné'}</Typography></>}
        <Button disabled={profile.isFetching} sx={{ ml: 'auto' }} onClick={() => void profile.refetch()}>Actualiser</Button>
      </Box>
      {profile.isPending ? <CircularProgress size={28} /> : profile.error ? <Alert severity="error">{profile.error.message}</Alert> : data && <>
        <Alert severity="info" sx={{ mb: 2 }}>Présences : fiches enregistrées avec un mois explicite, sans validation de complétude. Signatures : envois traçables depuis cette version ; l’absence de reçu ne prouve pas l’absence d’envoi. La signature des deux parties n’est pas synchronisée avec ClawShow.</Alert>
        {!!data.legacy_unbound_sheets && <Alert severity="warning" sx={{ mb: 2 }}>{data.legacy_unbound_sheets} ancienne(s) fiche(s) sans rattachement mensuel : non réparties entre les mois.</Alert>}
        {data.attendance_incomplete && <Alert severity="warning" sx={{ mb: 2 }}>Certaines anciennes configurations ne peuvent pas être lues. Le relevé des présences peut être incomplet.</Alert>}
        {!data.signatures_available && <Alert severity="warning" sx={{ mb: 2 }}>Historique des envois indisponible : aucun statut de signature ne peut être déduit.</Alert>}
        <TableContainer><Table size="small" aria-label={`Synthèse mensuelle ${year}`}><TableHead><TableRow>
          {['Mois', 'Présences', 'Documents', 'Signatures', 'Congés'].map(label => <TableCell key={label}>{label}</TableCell>)}
        </TableRow></TableHead><TableBody>{data.months.map(item => <TableRow key={item.month} selected={item.month === month}>
          <TableCell><Button aria-pressed={item.month === month} onClick={() => setMonth(item.month)}>{monthName(item.month)}</Button></TableCell>
          <TableCell>{item.attendance_sheets ? `${item.attendance_sheets} fiche(s) enregistrée(s)` : 'Aucune fiche liée'}</TableCell>
          <TableCell>{item.documents} · {item.documents - item.archived_documents} à traiter / {item.archived_documents} archivé(s)</TableCell>
          <TableCell>{item.signature_requests.length ? `${item.signature_requests.length} envoi(s) enregistré(s)` : data.signatures_available ? 'Aucun envoi traçable' : 'Indisponible'}</TableCell>
          <TableCell>{item.leaves.filter(leave => leave.status === 'pending').length} en attente · {item.leaves.filter(leave => leave.status === 'approved').length} approuvé(s)</TableCell>
        </TableRow>)}</TableBody></Table></TableContainer>
        {detail && <Paper variant="outlined" sx={{ p: 2, mt: 2 }}>
          <Typography variant="h6">Détail · {monthName(month)} {year}</Typography>
          <Typography variant="body2" sx={{ mt: 1 }}>{detail.attendance_sheets} fiche(s) liée(s) · {detail.documents} document(s), dont {detail.archived_documents} archivé(s).</Typography>
          <Typography sx={{ mt: 2, mb: 1, fontWeight: 500 }}>Demandes de congé touchant ce mois</Typography>
          {!detail.leaves.length && <Typography variant="body2" color="text.secondary">Aucune demande enregistrée.</Typography>}
          {detail.leaves.map(leave => <Box key={leave.id} sx={{ display: 'flex', flexWrap: 'wrap', alignItems: 'center', gap: 1, mb: 1 }}>
            <Typography variant="body2">{new Date(`${leave.start_date}T12:00:00`).toLocaleDateString('fr-FR')} ({slotLabels[leave.start_slot]}) → {new Date(`${leave.end_date}T12:00:00`).toLocaleDateString('fr-FR')} ({slotLabels[leave.end_slot]})</Typography>
            <Chip size="small" label={statusLabels[leave.status] || leave.status} variant="outlined" />
          </Box>)}
          <Typography sx={{ mt: 2, mb: 1, fontWeight: 500 }}>Envois pour signature</Typography>
          {!detail.signature_requests.length && <Typography variant="body2" color="text.secondary">Aucun reçu d’envoi lié à ce mois. Consultez ClawShow pour les demandes antérieures.</Typography>}
          {detail.signature_requests.map((item, index) => <Typography key={`${item.document_id}-${index}`} variant="body2" sx={{ overflowWrap: 'anywhere', mb: 1 }}>{item.document_id} · {item.mode} · {item.created_at} UTC · Enregistré chez le prestataire ; signature finale à vérifier dans ClawShow.</Typography>)}
        </Paper>}
      </>}
    </DialogContent>
    <DialogActions><Button onClick={onClose}>Fermer</Button></DialogActions>
  </Dialog>
}
