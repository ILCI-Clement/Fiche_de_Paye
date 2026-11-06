import { useState } from 'react'
import { Box, Button, Checkbox, Dialog, DialogActions, DialogContent, DialogContentText, DialogTitle, FormControlLabel } from '@mui/material'
import { withSystemAccess } from './accountRoles'

type SystemRole = 'Admin' | 'SuperAdmin'

export function SystemAccessFields({ roles, onChange, disabled, allowSuperAdmin = false, preserveAdmin = false, preserveSuperAdmin = false }: {
  roles: string[]; onChange: (roles: string[]) => void; disabled: boolean;
  allowSuperAdmin?: boolean; preserveAdmin?: boolean; preserveSuperAdmin?: boolean;
}) {
  const [pending, setPending] = useState<{ role: SystemRole; checked: boolean } | null>(null)
  const label = pending?.role === 'SuperAdmin' ? 'Super Admin' : 'Admin'
  return <>
    <Box role="group" aria-label="Droits système" sx={{ display: 'flex', alignItems: 'center', flexWrap: 'wrap', justifyContent: 'flex-end', gap: 1 }}>
      <FormControlLabel sx={{ m: 0 }} label="Admin" control={<Checkbox size="small"
        checked={roles.includes('Admin') || roles.includes('SuperAdmin')}
        disabled={disabled || preserveAdmin || roles.includes('SuperAdmin')}
        onChange={e => setPending({ role: 'Admin', checked: e.target.checked })} />} />
      {(allowSuperAdmin || roles.includes('SuperAdmin')) && <FormControlLabel sx={{ m: 0 }} label="Super Admin" control={<Checkbox size="small"
        checked={roles.includes('SuperAdmin')}
        disabled={disabled || !allowSuperAdmin || preserveSuperAdmin}
        onChange={e => setPending({ role: 'SuperAdmin', checked: e.target.checked })} />} />}
    </Box>
    <Dialog open={pending !== null} onClose={() => setPending(null)} maxWidth="xs" fullWidth aria-labelledby="system-access-confirm-title">
      <DialogTitle id="system-access-confirm-title">{pending?.checked ? `Accorder les droits ${label} ?` : `Retirer les droits ${label} ?`}</DialogTitle>
      <DialogContent><DialogContentText>
        {pending?.checked ? (pending.role === 'SuperAdmin' ? 'Accès complet au système, y compris les suppressions définitives.' : 'Accès à la gestion des comptes et des données du personnel.') : `Ce compte perdra les droits ${label}.`}
        {' '}Enregistrez ensuite pour appliquer ce choix.
      </DialogContentText></DialogContent>
      <DialogActions><Button onClick={() => setPending(null)}>Annuler</Button><Button variant="contained" disabled={disabled} onClick={() => {
        if (pending) onChange(withSystemAccess(roles, pending.role, pending.checked))
        setPending(null)
      }}>Confirmer</Button></DialogActions>
    </Dialog>
  </>
}
