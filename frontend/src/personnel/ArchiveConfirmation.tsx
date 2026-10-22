import { Box, Button, Dialog, DialogActions, DialogContent, DialogTitle, Typography } from '@mui/material'
import { personnelApi } from './api'
import type { Employee } from './api'
import { useRegistryAction } from './useRegistryAction'
import { WriteError } from './WriteError'

export function ArchiveConfirmation({ token, employee, onClose, onSaved }: { token: string; employee: Employee; onClose: () => void; onSaved: (message: string) => void }) {
  const action = useRegistryAction()
  async function save() {
    const ok = await action.run(() => personnelApi.archive(token, employee.employee_id, !employee.is_archived, employee.version))
    if (ok) { onSaved(employee.is_archived ? 'Fiche restaurée dans le registre actif.' : 'Fiche archivée ; les informations et l’historique sont conservés.'); onClose() }
  }
  return <Dialog open fullWidth maxWidth="sm" aria-labelledby="personnel-archive-title" onClose={() => { if (!action.busy) onClose() }}>
    <DialogTitle id="personnel-archive-title">{employee.is_archived ? 'Restaurer cette fiche ?' : 'Archiver cette fiche ?'}</DialogTitle>
    <DialogContent>
      <WriteError error={action.error} conflict={action.conflict} />
      <Typography sx={{ fontWeight: 500, mb: 1 }}>{employee.full_name}</Typography>
      <Typography variant="body2" color="text.secondary">{employee.is_archived ? 'La fiche sera à nouveau visible parmi les inscriptions actives.' : 'La fiche sera rangée dans les archives. Ses renseignements, périodes d’emploi et corrections resteront consultables.'}</Typography>
      <Box sx={{ mt: 2 }}><Typography variant="body2">Cette action ne modifie pas les dates d’entrée ou de sortie et ne change pas les droits du compte de connexion associé.</Typography></Box>
    </DialogContent>
    <DialogActions sx={{ p: 2 }}><Button disabled={action.busy} onClick={onClose}>Annuler</Button><Button variant="contained" disabled={action.busy || action.conflict} onClick={() => void save()}>{action.busy ? 'Enregistrement…' : employee.is_archived ? 'Restaurer' : 'Archiver'}</Button></DialogActions>
  </Dialog>
}
