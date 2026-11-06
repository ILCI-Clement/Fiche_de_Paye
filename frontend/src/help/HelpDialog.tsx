import { useId, useState } from 'react'
import { Box, Button, Dialog, DialogActions, DialogContent, DialogTitle, IconButton, Tab, Tabs, Typography } from '@mui/material'
import CloseRounded from '@mui/icons-material/CloseRounded'
import type { WorkspacePage } from '../workspacePages'
import { helpTopicsFor } from './helpTopics'

export function HelpDialog({ page, roles, onClose }: { page: WorkspacePage; roles: readonly string[]; onClose: () => void }) {
  const topics = helpTopicsFor({ role_tags: roles })
  const [selected, setSelected] = useState(() => topics.some(topic => topic.id === page) ? page : 'start')
  const topic = topics.find(item => item.id === selected) || topics[0]
  const id = useId()
  return <Dialog open fullWidth maxWidth="md" onClose={onClose} aria-labelledby={`${id}-title`}
    slotProps={{ paper: { sx: { height: 'min(640px, 88dvh)', maxHeight: '88dvh', m: 2, width: 'calc(100% - 32px)' } } }}>
    <DialogTitle id={`${id}-title`} sx={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', py: 1.5 }}>
      Guide d’utilisation<IconButton aria-label="Fermer le guide" onClick={onClose}><CloseRounded /></IconButton>
    </DialogTitle>
    <Box sx={{ display: 'flex', flex: 1, minHeight: 0, borderTop: 1, borderBottom: 1, borderColor: 'divider' }}>
      <Tabs value={topic.id} onChange={(_, value: string) => setSelected(value)} orientation="vertical" variant="scrollable" scrollButtons="auto"
        aria-label="Rubriques du guide" sx={{ width: { xs: 120, sm: 208 }, flexShrink: 0, borderRight: 1, borderColor: 'divider', py: 1,
          '& .MuiTab-root': { minHeight: 48, minWidth: 0, px: { xs: 1.5, sm: 2 }, fontSize: { xs: 12, sm: 14 }, textTransform: 'none', alignItems: 'flex-start', textAlign: 'left' },
          '& .Mui-selected': { bgcolor: 'action.selected' } }}>
        {topics.map(item => <Tab key={item.id} value={item.id} label={item.title} id={`${id}-${item.id}`} aria-controls={`${id}-${item.id}-panel`} />)}
      </Tabs>
    <DialogContent key={topic.id} role="tabpanel" id={`${id}-${topic.id}-panel`} aria-labelledby={`${id}-${topic.id}`} sx={{ px: { xs: 2, sm: 3 }, minHeight: 0, minWidth: 0, flex: 1, overflowWrap: 'break-word' }}>
      <Typography component="h2" variant="h6" sx={{ mb: 1 }}>{topic.title}</Typography>
      <Typography color="text.secondary" sx={{ mb: 2 }}>{topic.intro}</Typography>
      <Box component="ol" sx={{ pl: 2.5, m: 0, '& li': { pl: .5, mb: 1.5, lineHeight: 1.65 } }}>
        {topic.steps.map(step => <Typography component="li" variant="body2" key={step}>{step}</Typography>)}
      </Box>
      {topic.note && <Typography variant="body2" color="text.secondary" sx={{ mt: 2, pt: 2, borderTop: 1, borderColor: 'divider', lineHeight: 1.6 }}>{topic.note}</Typography>}
    </DialogContent>
    </Box>
    <DialogActions sx={{ px: 3, flexShrink: 0 }}><Button onClick={onClose}>Fermer</Button></DialogActions>
  </Dialog>
}
