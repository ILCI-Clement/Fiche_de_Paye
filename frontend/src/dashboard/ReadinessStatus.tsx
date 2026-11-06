import { useState } from 'react'
import { Box, Button, Dialog, DialogActions, DialogContent, DialogTitle, Divider, Tooltip, Typography } from '@mui/material'
import InfoOutlined from '@mui/icons-material/InfoOutlined'
import type { PersonReadiness } from './workCenterApi'
import { useWorkNavigation } from '../workspaceNavigation'

export function ReadinessStatus({ person }: { person: PersonReadiness }) {
  const [open, setOpen] = useState(false)
  const navigation = useWorkNavigation()
  if (!person.issues.length) return null
  return <>
    <Tooltip title={`${person.issues.length} information(s) à compléter`} enterDelay={500} arrow>
      <Button size="small" variant="outlined" color="inherit" onClick={() => setOpen(true)} aria-label={`État du dossier de ${person.display_name}`}
        startIcon={<InfoOutlined />} sx={{ whiteSpace: 'nowrap', minWidth: 0, minHeight: 26, height: 26, flexShrink: 0, px: .75, py: 0,
          fontSize: 11, fontWeight: 500, color: 'text.secondary', borderColor: 'divider', borderRadius: 1,
          '& .MuiButton-startIcon': { ml: 0, mr: .5 }, '& .MuiButton-startIcon > svg': { fontSize: 14 },
          '&:hover': { borderColor: 'text.secondary', bgcolor: 'action.hover' } }}>
        {person.issues.length}
      </Button>
    </Tooltip>
    <Dialog open={open} onClose={() => setOpen(false)} fullWidth maxWidth="sm">
      <DialogTitle>État du dossier · {person.display_name}</DialogTitle>
      <DialogContent>
        {person.issues.map(issue => <Box key={issue.code} sx={{ mb: 1 }}><Box sx={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 1, py: .75 }}>
          <Box><Typography variant="body2">{issue.label}</Typography>
          {!issue.destination && <Typography variant="caption" color="text.secondary" sx={{ display: 'block' }}>À faire vérifier par un administrateur ou le responsable direct.</Typography>}</Box>
          {issue.destination && navigation && <Button size="small" sx={{ flexShrink: 0 }} onClick={() => { setOpen(false); navigation.navigate({ ...issue.destination!, completion: issue.code }) }}>Compléter</Button>}
        </Box><Divider /></Box>)}
      </DialogContent>
      <DialogActions><Button onClick={() => setOpen(false)}>Fermer</Button></DialogActions>
    </Dialog>
  </>
}
