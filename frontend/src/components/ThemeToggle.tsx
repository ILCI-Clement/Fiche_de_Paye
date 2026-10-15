import { IconButton, Tooltip } from '@mui/material'
import DarkModeOutlined from '@mui/icons-material/DarkModeOutlined'
import LightModeOutlined from '@mui/icons-material/LightModeOutlined'
import { useThemeMode } from '../themeMode'

export function ThemeToggle() {
  const { mode, toggle } = useThemeMode()
  const label = mode === 'dark' ? 'Passer en mode clair (Light)' : 'Passer en mode sombre (Dark)'
  return <Tooltip title={label}><IconButton onClick={toggle} aria-label={label}>
    {mode === 'dark' ? <LightModeOutlined /> : <DarkModeOutlined />}
  </IconButton></Tooltip>
}
