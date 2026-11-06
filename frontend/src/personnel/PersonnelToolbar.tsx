import { useId, useState } from 'react'
import { Badge, Box, Button, IconButton, InputAdornment, MenuItem, Popover, TextField, Tooltip, Typography } from '@mui/material'
import SearchRounded from '@mui/icons-material/SearchRounded'
import PersonAddOutlined from '@mui/icons-material/PersonAddOutlined'
import FilterListRounded from '@mui/icons-material/FilterListRounded'
import RefreshRounded from '@mui/icons-material/RefreshRounded'
import FilterAltOffOutlined from '@mui/icons-material/FilterAltOffOutlined'
import InfoOutlined from '@mui/icons-material/InfoOutlined'
import CloseRounded from '@mui/icons-material/CloseRounded'
import { personnelTypes, statusLabels } from './api'
import type { PersonnelFilters } from './api'

interface Props {
  filters: PersonnelFilters; search: string; busy: boolean;
  onSearchChange: (value: string) => void; onSearch: () => void;
  onFilter: (value: Partial<PersonnelFilters>) => void; onReset: () => void;
  onRegister: () => void; onRefresh: () => void;
}

/** Rendered by WorkspaceToolbar in the active page's application header. */
export function PersonnelToolbar({ filters, search, busy, onSearchChange, onSearch, onFilter, onReset, onRegister, onRefresh }: Props) {
  const [anchor, setAnchor] = useState<HTMLButtonElement | null>(null)
  const filterId = useId()
  const filterCount = Number(!!filters.personnel_type) + Number(filters.status !== 'all') + Number(filters.archive !== 'active')
  const selected = !!filters.q || filterCount > 0
  return <Box sx={{ display: 'flex', alignItems: 'center', flexWrap: 'wrap', gap: .75, width: '100%', minWidth: 0,
    '& .MuiIconButton-root': { width: 40, height: 40 }, '& .MuiButton-root': { minHeight: 40, fontWeight: 500 },
  }}>
    <TextField select size="small" label="Dossier" value={filters.archive} disabled={busy} sx={{ width: 140 }} onChange={event => onFilter({ archive: event.target.value as PersonnelFilters['archive'] })}>
      <MenuItem value="active">Normal</MenuItem><MenuItem value="archived">Archivé</MenuItem><MenuItem value="all">Tous</MenuItem>
    </TextField>
    <Box component="form" onSubmit={event => { event.preventDefault(); onSearch() }} sx={{ flex: '1 1 200px', minWidth: 170, maxWidth: { xs: 'none', md: 340 } }}>
      <TextField placeholder="Nom, compte ou n°…" value={search} onChange={event => onSearchChange(event.target.value)} fullWidth
        sx={{ '& .MuiInputBase-root': { minHeight: 40, height: 40 }, '& .MuiInputBase-input': { py: 1, fontSize: 13 } }}
        slotProps={{ htmlInput: { maxLength: 200, 'aria-label': 'Rechercher dans le liste des utilisateurs' }, input: { endAdornment: <InputAdornment position="end"><Tooltip title="Rechercher"><span><IconButton type="submit" disabled={busy} aria-label="Lancer la recherche" edge="end"><SearchRounded fontSize="small" /></IconButton></span></Tooltip></InputAdornment> } }} />
    </Box>
    <Tooltip title={filterCount ? `Filtres (${filterCount})` : 'Filtrer le registre'}><IconButton aria-label="Filtrer le registre" aria-haspopup="dialog" aria-expanded={!!anchor} aria-controls={anchor ? filterId : undefined} onClick={event => setAnchor(event.currentTarget)} color={filterCount ? 'primary' : 'default'}><Badge badgeContent={filterCount} color="primary"><FilterListRounded fontSize="small" /></Badge></IconButton></Tooltip>
    <Tooltip title="Actualiser"><span><IconButton aria-label="Actualiser le registre" disabled={busy} onClick={onRefresh}><RefreshRounded fontSize="small" /></IconButton></span></Tooltip>
    {selected && <Tooltip title="Réinitialiser la recherche et les filtres"><IconButton aria-label="Réinitialiser la recherche et les filtres" onClick={onReset}><FilterAltOffOutlined fontSize="small" /></IconButton></Tooltip>}
    <Tooltip title="Une ligne par personne ; la fiche conserve toutes les périodes d’emploi. L’état correspond à la dernière période datée. L’archivage ne change pas l’accès au compte."><IconButton aria-label="Informations sur le registre"><InfoOutlined fontSize="small" /></IconButton></Tooltip>
    <Button className="workspace-primary-action" variant="contained" size="small" startIcon={<PersonAddOutlined fontSize="small" />} onClick={onRegister} sx={{ flexShrink: 0, whiteSpace: 'nowrap' }}>Enregistrer un salarié</Button>
    <Popover id={filterId} open={!!anchor} anchorEl={anchor} onClose={() => setAnchor(null)} anchorOrigin={{ vertical: 'bottom', horizontal: 'right' }} transformOrigin={{ vertical: 'top', horizontal: 'right' }}>
      <Box role="dialog" aria-label="Filtres du registre" sx={{ p: 2, width: 300, maxWidth: 'calc(100vw - 32px)' }}>
        <Box sx={{ display: 'flex', alignItems: 'center', mb: 2 }}><Typography sx={{ flex: 1, fontWeight: 500 }}>Filtres du registre</Typography><IconButton size="small" aria-label="Fermer les filtres" onClick={() => setAnchor(null)}><CloseRounded fontSize="small" /></IconButton></Box>
        <Box sx={{ display: 'flex', flexDirection: 'column', gap: 2 }}>
          <TextField select fullWidth label="Personnel" value={filters.personnel_type} onChange={event => onFilter({ personnel_type: event.target.value as PersonnelFilters['personnel_type'] })}>
            <MenuItem value="">Tous les types</MenuItem>{personnelTypes.filter(item => item.value).map(item => <MenuItem key={item.value} value={item.value}>{item.label}</MenuItem>)}
          </TextField>
          <TextField select fullWidth label="État d’emploi" value={filters.status} onChange={event => onFilter({ status: event.target.value as PersonnelFilters['status'] })}><MenuItem value="all">Tous les états</MenuItem>{Object.entries(statusLabels).map(([value, label]) => <MenuItem key={value} value={value}>{label}</MenuItem>)}</TextField>
          <Button disabled={!selected} onClick={onReset}>Réinitialiser les filtres et la recherche</Button>
        </Box>
      </Box>
    </Popover>
  </Box>
}
