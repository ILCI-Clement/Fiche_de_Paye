import { useState } from 'react'
import { useQuery } from '@tanstack/react-query'
import { Alert, Box, Button, Dialog, DialogActions, DialogContent, DialogTitle, LinearProgress, Typography } from '@mui/material'
import StorageOutlined from '@mui/icons-material/StorageOutlined'
import { request } from '../api'

type Volume = { labels: string[]; total_bytes: number; used_bytes: number; free_bytes: number; used_percent: number; level: 'normal' | 'warning' | 'critical' }
type Capacity = { volumes: Volume[]; unavailable: string[]; checked_at: string }
const gib = (bytes: number) => `${(bytes / 1024 ** 3).toLocaleString('fr-FR', { maximumFractionDigits: 1 })} GiB`
const color = (level: Volume['level']) => level === 'critical' ? 'error' : level === 'warning' ? 'warning' : 'primary'

export function StoragePanel({ token, local }: { token: string; local: boolean }) {
  const [open, setOpen] = useState(false)
  const query = useQuery({ queryKey: ['system-storage', token],
    queryFn: ({ signal }) => request<Capacity>('/system/storage', token, { signal, cache: 'no-store' }),
    refetchInterval: 60000, staleTime: 30000, retry: false })
  const data = query.isError ? undefined : query.data
  const worst = data?.volumes.reduce<Volume | undefined>((current, volume) => !current ||
    ({ normal: 0, warning: 1, critical: 2 }[volume.level] > { normal: 0, warning: 1, critical: 2 }[current.level]) ||
    (volume.level === current.level && volume.used_percent > current.used_percent) ? volume : current, undefined)
  return <>
    <Button onClick={() => setOpen(true)} startIcon={<StorageOutlined fontSize="small" />} size="small"
      sx={{ width: '100%', justifyContent: 'space-between', px: 1.5, color: worst && worst.level !== 'normal' ? `${color(worst.level)}.main` : 'var(--sidebar-muted)', fontSize: 11 }}>
      {local ? 'Stockage local' : 'Stockage serveur'}
      <span>{query.isError || !worst ? (query.isPending ? '…' : 'Indisponible') : `${worst.used_percent.toLocaleString('fr-FR')} %${data?.unavailable.length ? ' · !' : ''}`}</span>
    </Button>
    <Dialog open={open} onClose={() => setOpen(false)} fullWidth maxWidth="sm">
      <DialogTitle>{local ? 'Stockage de cette machine' : 'Stockage du serveur'}</DialogTitle>
      <DialogContent>
        {query.isError && <Alert severity="warning">Capacité indisponible. Réessayez.</Alert>}
        {query.isPending && <LinearProgress />}
        {data?.volumes.map((volume, index) => <Box key={index} sx={{ mb: 2.5 }}>
          <Typography sx={{ mb: .5, fontWeight: 500 }}>{volume.labels.join(' / ')}</Typography>
          <Box sx={{ display: 'flex', justifyContent: 'space-between', gap: 1, mb: 1 }}>
            <Typography variant="body2">{gib(volume.used_bytes)} / {gib(volume.total_bytes)}</Typography>
            <Typography variant="body2">{gib(volume.free_bytes)} libres</Typography>
          </Box>
          <LinearProgress variant="determinate" value={volume.used_percent} color={color(volume.level)} sx={{ height: 6, borderRadius: 1 }} />
          {volume.level !== 'normal' && <Alert severity={volume.level === 'critical' ? 'error' : 'warning'} sx={{ mt: 1 }}>
            {volume.level === 'critical' ? 'Espace disque critique : prévoir une intervention.' : 'Le disque approche de sa capacité maximale.'}
          </Alert>}
        </Box>)}
        {!!data?.unavailable.length && <Alert severity="warning">Mesure indisponible : {data.unavailable.join(', ')}.</Alert>}
        <Typography variant="caption" color="text.secondary" component="p">Capacité des disques hébergeant l’application et ses fichiers, y compris les autres données présentes. Les disques partagés ne sont comptés qu’une fois. Le stockage de la base de données peut être distinct.</Typography>
        {data && <Typography variant="caption" color="text.secondary">Actualisé le {new Date(data.checked_at).toLocaleString('fr-FR')}. Actualisation chaque minute lorsque cette fenêtre est active.</Typography>}
      </DialogContent>
      <DialogActions><Button disabled={query.isFetching} onClick={() => void query.refetch()}>Actualiser</Button><Button onClick={() => setOpen(false)}>Fermer</Button></DialogActions>
    </Dialog>
  </>
}
