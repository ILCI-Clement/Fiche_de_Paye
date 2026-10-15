import { useId, useState } from 'react'
import { Button, ListItemIcon, ListItemText, Menu, MenuItem } from '@mui/material'
import MoreHorizRounded from '@mui/icons-material/MoreHorizRounded'
import DownloadRounded from '@mui/icons-material/DownloadRounded'
import PictureAsPdfOutlined from '@mui/icons-material/PictureAsPdfOutlined'

export function FicheExportMenu({ intern, busy, onExport }: { intern: boolean; busy: boolean; onExport: (format: 'original' | 'pdf') => void }) {
  const [anchor, setAnchor] = useState<HTMLElement | null>(null)
  const id = useId()
  const open = Boolean(anchor)
  function download(format: 'original' | 'pdf') { setAnchor(null); onExport(format) }
  return <>
    <Button id={`${id}-button`} color="inherit" startIcon={<MoreHorizRounded />} disabled={busy}
      aria-haspopup="menu" aria-expanded={open ? 'true' : undefined} aria-controls={open ? `${id}-menu` : undefined}
      onClick={event => setAnchor(event.currentTarget)}>More</Button>
    <Menu id={`${id}-menu`} anchorEl={anchor} open={open} onClose={() => setAnchor(null)} slotProps={{ list: { 'aria-labelledby': `${id}-button` } }}>
      <MenuItem disabled={busy} onClick={() => download('original')}><ListItemIcon><DownloadRounded fontSize="small" /></ListItemIcon><ListItemText>{intern ? 'Générer le Word' : 'Générer l’Excel'}</ListItemText></MenuItem>
      <MenuItem disabled={busy} onClick={() => download('pdf')}><ListItemIcon><PictureAsPdfOutlined fontSize="small" /></ListItemIcon><ListItemText>Télécharger le PDF</ListItemText></MenuItem>
    </Menu>
  </>
}
