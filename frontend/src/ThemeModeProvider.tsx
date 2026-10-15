import { useMemo, useState } from 'react'
import type { PropsWithChildren } from 'react'
import { CssBaseline, ThemeProvider } from '@mui/material'
import { buildTheme } from './theme'
import { ThemeModeContext } from './themeMode'
import type { Mode } from './themeMode'

const STORAGE_KEY = 'presence-color-mode'

function initialMode(): Mode {
  try { return localStorage.getItem(STORAGE_KEY) === 'light' ? 'light' : 'dark' }
  catch { return 'dark' }
}

export function ThemeModeProvider({ children }: PropsWithChildren) {
  const [mode, setMode] = useState<Mode>(initialMode)
  const theme = useMemo(() => buildTheme(mode), [mode])
  const value = useMemo(() => ({ mode, toggle: () => {
    const next = mode === 'dark' ? 'light' : 'dark'
    try { localStorage.setItem(STORAGE_KEY, next) } catch { /* Still switch when storage is unavailable. */ }
    setMode(next)
  } }), [mode])
  return <ThemeModeContext.Provider value={value}><ThemeProvider theme={theme}>
    <CssBaseline enableColorScheme />{children}
  </ThemeProvider></ThemeModeContext.Provider>
}

