import { StrictMode } from 'react'
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
      <QueryClientProvider client={queryClient}><App /></QueryClientProvider>
    </ThemeModeProvider>
  </StrictMode>,
)
