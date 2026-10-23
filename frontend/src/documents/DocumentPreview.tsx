import { lazy, Suspense, useEffect, useState } from 'react'
import { Alert, Box, Button, CircularProgress, Dialog, DialogActions, DialogContent, DialogTitle, Typography, useMediaQuery } from '@mui/material'
import { useTheme } from '@mui/material/styles'
import { documentsApi } from './api'
import type { Receipt } from './api'
import { downloadFile } from '../download'

const PdfPreview = lazy(() => import('./PdfPreview'))
interface PreviewFile { blob: Blob; filename: string; url: string; pdf: boolean; imageSource?: string }

export function DocumentPreview({ token, receipt, onClose, loadFile }: { token: string; receipt: Pick<Receipt, 'id' | 'original_filename' | 'document_number'>; onClose: () => void; loadFile?: (signal: AbortSignal) => Promise<{ blob: Blob; filename: string }> }) {
  const fullScreen = useMediaQuery(useTheme().breakpoints.down('sm'))
  const [file, setFile] = useState<PreviewFile | null>(null)
  const [error, setError] = useState('')
  const [attempt, setAttempt] = useState(0)
  useEffect(() => {
    const controller = new AbortController()
    let url: string | undefined
    void (loadFile ? loadFile(controller.signal) : documentsApi.file(token, receipt.id, controller.signal)).then(async result => {
      if (controller.signal.aborted) return
      const type = result.blob.type.split(';')[0].toLowerCase()
      if (!['application/pdf', 'image/jpeg', 'image/png'].includes(type)) {
        throw new Error('Ce format ne permet pas un aperçu. Utilisez le téléchargement de la carte du document.')
      }
      const imageSource = type === 'application/pdf' ? undefined : await new Promise<string>((resolve, reject) => {
        const reader = new FileReader()
        reader.onload = () => resolve(String(reader.result))
        reader.onerror = () => reject(new Error('Impossible de lire cette image.'))
        reader.readAsDataURL(result.blob)
      })
      if (controller.signal.aborted) return
      url = URL.createObjectURL(result.blob)
      setFile({ ...result, url, pdf: type === 'application/pdf', imageSource })
    }).catch(reason => {
      if (!controller.signal.aborted) setError(reason instanceof Error ? reason.message : 'Impossible de charger le document.')
    })
    return () => { controller.abort(); if (url) URL.revokeObjectURL(url) }
  }, [token, receipt.id, attempt, loadFile])
  return <Dialog open fullWidth maxWidth="lg" fullScreen={fullScreen} onClose={onClose} aria-labelledby="document-preview-title">
    <DialogTitle id="document-preview-title" sx={{ overflowWrap: 'anywhere' }}>
      Aperçu · {receipt.original_filename}
      {receipt.document_number && <Typography component="span" variant="body2" color="text.secondary" sx={{ display: 'block' }}>N° {receipt.document_number}</Typography>}
    </DialogTitle>
    <DialogContent dividers sx={{ p: { xs: 1, sm: 2 }, minHeight: 240, display: 'flex', flexDirection: 'column' }}>
      {error ? <Alert severity="error" action={<Button color="inherit" onClick={() => { setFile(null); setError(''); setAttempt(value => value + 1) }}>Réessayer</Button>}>{error}</Alert>
        : !file ? <Box role="status" sx={{ m: 'auto', p: 4, textAlign: 'center' }}><CircularProgress size={28} /><Typography sx={{ mt: 2 }}>Chargement du document…</Typography></Box>
        : file.pdf ? <Suspense fallback={<CircularProgress size={28} sx={{ m: 'auto' }} />}><PdfPreview blob={file.blob} /></Suspense>
        : <Box component="img" src={file.imageSource} alt={receipt.original_filename} onError={() => setError('Impossible d’afficher cette image. Vous pouvez la télécharger.')}
          sx={{ display: 'block', maxWidth: '100%', maxHeight: fullScreen ? '100%' : '65vh', objectFit: 'contain', m: 'auto', minHeight: 0 }} />}
    </DialogContent>
    <DialogActions sx={{ px: 2, py: 1.5, flexWrap: 'wrap', gap: 1 }}>
      {file && <>
        <Button component="a" href={file.url} target="_blank" rel="noopener noreferrer">Ouvrir dans un onglet</Button>
        <Button onClick={() => downloadFile(file)}>Télécharger</Button>
      </>}
      <Button variant="outlined" onClick={onClose}>Fermer</Button>
    </DialogActions>
  </Dialog>
}
