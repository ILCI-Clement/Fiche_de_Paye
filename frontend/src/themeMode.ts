import { createContext, useContext } from 'react'

export type Mode = 'light' | 'dark'
export const ThemeModeContext = createContext<{ mode: Mode; toggle: () => void } | null>(null)

export function useThemeMode() {
  const context = useContext(ThemeModeContext)
  if (!context) throw new Error('ThemeModeProvider is required.')
  return context
}
