import { personDisplayName } from '../personDisplayName'
import { useCallback, useState } from 'react'
import { Alert, Box, Button, CircularProgress, Dialog, DialogActions, DialogContent, DialogTitle, IconButton, MenuItem, Paper, TextField, Tooltip, Typography } from '@mui/material'
import HelpOutlineRounded from '@mui/icons-material/HelpOutlineRounded'
import { useQuery } from '@tanstack/react-query'
import { request } from '../api'
import { DepartmentBadges } from '../components/DepartmentBadge'
import { AccountPersonnelSection } from '../profile/AccountPersonnelSection'
import { DocumentPreview } from '../documents/DocumentPreview'
import { periodLabel } from '../documents/api'
import { signedFichesApi } from '../attendance/signedFichesApi'
import { ProfileMonth } from './ProfileMonth'
import type { MonthlyProfile } from './ProfileMonth'

interface Profile {
  username: string; name: string; display_name?: string; full_name?: string | null; manager_name?: string; groups: string[]; manager: string | null; year: number; months: MonthlyProfile[];
  legacy_unbound_sheets: number; attendance_incomplete: boolean; signatures_available: boolean;
  can_view_personnel_data?: boolean;
  leave_totals?: { approved_days: number; incomplete: boolean; unverified_requests: number };
}

function ProfileSignedPreview({ token, archiveId, filename, onClose }: { token: string; archiveId: string; filename: string; onClose: () => void }) {
  const loadFile = useCallback((signal: AbortSignal) => signedFichesApi.file(token, archiveId, signal), [token, archiveId])
  return <DocumentPreview token={token} receipt={{ id: 0, original_filename: filename }} loadFile={loadFile} onClose={onClose} />
}

export function PersonProfile({ username, token, viewer, onClose }: { username: string; token: string; viewer: string; onClose: () => void }) {
  const [today] = useState(() => new Date())
  const [year, setYear] = useState(() => today.getFullYear())
  const [month, setMonth] = useState(() => today.getMonth() + 1)
  const [preview, setPreview] = useState<{ id: string; month: number } | null>(null)
  const profile = useQuery({ queryKey: ['person-profile', viewer, username, year], queryFn: ({ signal }) => request<Profile>(`/person-profiles/${encodeURIComponent(username)}?year=${year}`, token, { signal }), retry: false })
  const data = profile.data
  const detail = data?.months.find(item => item.month === month)
  const name = personDisplayName(data ? { full_name: data.full_name, display_name: data.display_name || data.name } : null)
  return <Dialog open fullWidth maxWidth="md" onClose={onClose} aria-labelledby="person-profile-title"
    slotProps={{ paper: { sx: { height: 'min(560px, calc(100dvh - 48px))' } } }}>
    <DialogTitle id="person-profile-title">Profil · {name}</DialogTitle>
    <DialogContent dividers sx={{ minWidth: 0 }}>
      <Box sx={{ display: 'flex', gap: 1.5, flexWrap: 'wrap', alignItems: 'center', mb: 2 }}>
        <TextField select size="small" label="Année" value={year} onChange={event => { setYear(Number(event.target.value)); setPreview(null) }} sx={{ width: 120 }}>
          {Array.from({ length: 101 }, (_, i) => 2000 + i).reverse().map(value => <MenuItem key={value} value={value}>{value}</MenuItem>)}
        </TextField>
        <TextField select size="small" label="Mois" value={month} onChange={event => { setMonth(Number(event.target.value)); setPreview(null) }} sx={{ width: 150 }}>
          {Array.from({ length: 12 }, (_, index) => <MenuItem key={index + 1} value={index + 1} sx={{ textTransform: 'capitalize' }}>{new Intl.DateTimeFormat('fr-FR', { month: 'long' }).format(new Date(2026, index, 1))}</MenuItem>)}
        </TextField>
        <Tooltip title="Les détails concernent le mois sélectionné. Les anciens envois sans suivi restent à vérifier dans ClawShow ; l’absence de reçu ne prouve pas l’absence d’envoi.">
          <IconButton size="small" aria-label="À propos du profil mensuel"><HelpOutlineRounded sx={{ fontSize: 17 }} /></IconButton>
        </Tooltip>
        {data && <><DepartmentBadges names={data.groups} /><Typography variant="body2" color="text.secondary">Responsable direct : {data.manager ? personDisplayName({ display_name: data.manager_name }) : 'Non renseigné'}</Typography></>}
        <Button size="small" disabled={profile.isFetching} sx={{ ml: 'auto' }} onClick={() => void profile.refetch()}>Actualiser</Button>
      </Box>
      {profile.isPending ? <CircularProgress size={28} /> : profile.error ? <Alert severity="error">{profile.error.message}</Alert> : data && <>
        {data.can_view_personnel_data && <AccountPersonnelSection token={token} username={username} viewer={viewer} />}
        {!!data.legacy_unbound_sheets && <Alert severity="warning" sx={{ mb: 2 }}>{data.legacy_unbound_sheets} ancienne(s) fiche(s) sans rattachement mensuel : non réparties entre les mois.</Alert>}
        {data.attendance_incomplete && <Alert severity="warning" sx={{ mb: 2 }}>Certaines anciennes configurations ne peuvent pas être lues. Le relevé des présences peut être incomplet.</Alert>}
        {!data.signatures_available && <Alert severity="warning" sx={{ mb: 2 }}>Historique des envois indisponible : aucun statut de signature ne peut être déduit.</Alert>}
        {data.leave_totals && <Tooltip title={data.leave_totals.incomplete
          ? `${data.leave_totals.unverified_requests} ancienne(s) demande(s) sans durée vérifiable : total partiel.`
          : 'Demandes approuvées, dates futures incluses. Demi-journée = 0,5 jour. Repos, jours fériés et doublons exclus.'}>
          <Paper variant="outlined" sx={{ display: 'flex', alignItems: 'center', gap: 1, px: 1.5, py: 1, mb: 1.5 }}>
            <Typography variant="body2" sx={{ flex: 1 }}>Congés approuvés · {year}{data.leave_totals.incomplete ? ' · À vérifier' : ''}</Typography>
            <Typography sx={{ fontWeight: 600 }}>{new Intl.NumberFormat('fr-FR', { maximumFractionDigits: 1 }).format(data.leave_totals.approved_days)} jour(s){data.leave_totals.incomplete ? ' confirmés' : ''}</Typography>
          </Paper>
        </Tooltip>}
        <Box component="section" aria-label={`Détails · ${periodLabel(month, year)}`} sx={{ minWidth: 0 }}>
          {detail ? <ProfileMonth key={`${year}-${month}`} item={detail} year={year} current={year === today.getFullYear() && month === today.getMonth() + 1} signaturesAvailable={data.signatures_available} onPreview={(id, selectedMonth) => setPreview({ id, month: selectedMonth })} /> : <Typography color="text.secondary">Aucune donnée disponible pour ce mois.</Typography>}
        </Box>
      </>}
    </DialogContent>
    <DialogActions><Button onClick={onClose}>Fermer</Button></DialogActions>
    {preview && <ProfileSignedPreview token={token} archiveId={preview.id} filename={`${name} · ${periodLabel(preview.month, year)} · signé.pdf`} onClose={() => setPreview(null)} />}
  </Dialog>
}
