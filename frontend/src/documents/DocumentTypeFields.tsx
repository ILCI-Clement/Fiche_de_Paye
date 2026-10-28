import { MenuItem, TextField } from '@mui/material'
import type { DocumentCategory } from './api'
import { DocumentTypeBadge } from './DocumentTypeBadge'

export function DocumentTypeFields({ categories, type, other, onType, onOther, disabled }: {
  categories: DocumentCategory[]; type: string; other: string; onType: (value: string) => void; onOther: (value: string) => void; disabled: boolean
}) {
  return <>
    <TextField select required label="Type de document" value={type} disabled={disabled} onChange={event => onType(event.target.value)} slotProps={{ select: {
      renderValue: selected => { const category = categories.find(item => item.code === selected); return category ? <DocumentTypeBadge code={category.code} label={category.label} inactive={!category.is_active} /> : 'Choisir un type' },
      MenuProps: { slotProps: { paper: { sx: { maxHeight: 300, maxWidth: 'calc(100vw - 32px)' } } } },
    } }} sx={{ minWidth: 180, maxWidth: '100%', flex: '1 1 180px' }}>
      {!categories.some(category => category.code === type) && <MenuItem value={type} disabled>Choisir un type</MenuItem>}
      {categories.filter(category => category.is_active || category.code === type).map(category => <MenuItem key={category.code} value={category.code} disabled={!category.is_active} sx={{ whiteSpace: 'normal', overflowWrap: 'anywhere' }}>
        <DocumentTypeBadge code={category.code} label={`${category.label}${!category.is_active ? ' (désactivé)' : ''}`} inactive={!category.is_active} />
      </MenuItem>)}
    </TextField>
    {type === 'autre' && <TextField required label="Précisez le type" value={other} onChange={event => onOther(event.target.value)} disabled={disabled} slotProps={{ htmlInput: { maxLength: 120 } }} sx={{ flex: '1 1 180px' }} />}
  </>
}
