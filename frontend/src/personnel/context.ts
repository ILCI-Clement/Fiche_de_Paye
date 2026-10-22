import { createContext, useContext } from 'react'

export const RegistryViewerContext = createContext('')
export function useRegistryViewer() { return useContext(RegistryViewerContext) }
