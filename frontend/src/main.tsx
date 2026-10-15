import { StrictMode } from 'react'
import { LocalizationProvider } from '@mui/x-date-pickers/LocalizationProvider'
import { AdapterDayjs } from '@mui/x-date-pickers/AdapterDayjs'
import { frFR } from '@mui/x-date-pickers/locales'
import 'dayjs/locale/fr'
import { createRoot } from 'react-dom/client'
import './index.css'
import App from './App.tsx'
import { QueryClient, QueryClientProvider } from '@tanstack/react-query'
import { ThemeModeProvider } from './ThemeModeProvider'

const queryClient = new QueryClient({ defaultOptions: {
  queries: { retry: false, staleTime: 15000, refetchOnWindowFocus: true },
  mutations: { retry: false },
} })

createRoot(document.getElementById('root')!).render(
  <StrictMode>
    <ThemeModeProvider>
      <LocalizationProvider dateAdapter={AdapterDayjs} adapterLocale="fr" localeText={frFR.components.MuiLocalizationProvider.defaultProps.localeText}>
        <QueryClientProvider client={queryClient}><App /></QueryClientProvider>
      </LocalizationProvider>
    </ThemeModeProvider>
  </StrictMode>,
)
