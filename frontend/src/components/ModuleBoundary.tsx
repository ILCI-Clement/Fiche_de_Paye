import { Component, lazy, Suspense } from 'react'
import type { ComponentType, ReactNode } from 'react'
import { Alert, Box, Button, CircularProgress } from '@mui/material'

class ModuleError extends Component<{ children: ReactNode; retry: () => void }, { failed: boolean }> {
  state = { failed: false }
  static getDerivedStateFromError() { return { failed: true } }
  render() {
    return this.state.failed ? <Alert severity="error">Impossible d’ouvrir ce module. Vérifiez la connexion puis réessayez.
      <Button onClick={this.props.retry}>Réessayer</Button>
    </Alert> : this.props.children
  }
}

/** Recreate the lazy import after rejection; keep sibling modules and their drafts mounted. */
export function recoverableModule<P extends object>(load: () => Promise<{ default: ComponentType<P> }>): ComponentType<P> {
  return class RecoverableModule extends Component<P, { revision: number }> {
    state = { revision: 0 }
    Module = lazy(load)
    retry = () => { this.Module = lazy(load); this.setState(previous => ({ revision: previous.revision + 1 })) }
    render() {
      const Module = this.Module
      return <ModuleError key={this.state.revision} retry={this.retry}>
        <Suspense fallback={<Box className="workspace"><CircularProgress size={26} /></Box>}><Module {...this.props} /></Suspense>
      </ModuleError>
    }
  }
}

