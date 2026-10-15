import { useEffect, useRef, useState } from 'react'
import { Alert, Box, Button, CircularProgress, Typography } from '@mui/material'
import * as pdfjs from 'pdfjs-dist'
import workerUrl from 'pdfjs-dist/build/pdf.worker.min.mjs?worker&url'
import type { PDFDocumentProxy, RenderTask } from 'pdfjs-dist'

pdfjs.GlobalWorkerOptions.workerSrc = workerUrl
const assets = `${import.meta.env.BASE_URL}pdfjs/${pdfjs.version}/`

export default function PdfPreview({ blob }: { blob: Blob }) {
  const [document, setDocument] = useState<PDFDocumentProxy | null>(null)
  const [page, setPage] = useState(1)
  const [zoom, setZoom] = useState(1)
  const [width, setWidth] = useState(800)
  const [error, setError] = useState('')
  const container = useRef<HTMLDivElement>(null)
  const canvas = useRef<HTMLCanvasElement>(null)
  useEffect(() => {
    let cancelled = false
    let loading: ReturnType<typeof pdfjs.getDocument> | undefined
    void blob.arrayBuffer().then(data => {
      if (cancelled) return
      loading = pdfjs.getDocument({ data: new Uint8Array(data), cMapUrl: `${assets}cmaps/`, cMapPacked: true,
        standardFontDataUrl: `${assets}standard_fonts/`, wasmUrl: `${assets}wasm/`, iccUrl: `${assets}iccs/` })
      return loading.promise.then(pdf => { if (!cancelled) setDocument(pdf) })
    }).catch(() => { if (!cancelled) setError('Ce PDF ne peut pas être affiché. Vous pouvez le télécharger ou l’ouvrir dans un onglet.') })
    return () => { cancelled = true; void loading?.destroy() }
  }, [blob])
  useEffect(() => {
    const target = container.current
    if (!target) return
    const observer = new ResizeObserver(entries => setWidth(Math.max(160, Math.floor(entries[0].contentRect.width))))
    observer.observe(target)
    return () => observer.disconnect()
  }, [])
  useEffect(() => {
    if (!document || !canvas.current) return
    const target = canvas.current
    let cancelled = false
    let render: RenderTask | undefined
    void document.getPage(page).then(pdfPage => {
      if (cancelled) return
      const scale = Math.max(100, width - 16) / pdfPage.getViewport({ scale: 1 }).width * zoom
      const ratio = Math.min(window.devicePixelRatio || 1, 2)
      const viewport = pdfPage.getViewport({ scale: scale * ratio })
      target.width = Math.ceil(viewport.width); target.height = Math.ceil(viewport.height)
      target.style.width = `${viewport.width / ratio}px`; target.style.height = `${viewport.height / ratio}px`
      render = pdfPage.render({ canvas: target, viewport, background: '#ffffff' })
      return render.promise
    }).catch(reason => { if (!cancelled && reason?.name !== 'RenderingCancelledException') setError('Impossible d’afficher cette page. Le téléchargement reste disponible.') })
    return () => { cancelled = true; render?.cancel() }
  }, [document, page, zoom, width])
  return <Box sx={{ minWidth: 0 }}>
    {error && <Alert severity="error" sx={{ mb: 1 }}>{error}</Alert>}
    <Box sx={{ display: 'flex', alignItems: 'center', flexWrap: 'wrap', gap: 1, mb: 1 }}>
      <Button disabled={!document || page <= 1} onClick={() => setPage(value => value - 1)}>Précédente</Button>
      <Typography variant="body2">Page {page} / {document?.numPages || '…'}</Typography>
      <Button disabled={!document || page >= document.numPages} onClick={() => setPage(value => value + 1)}>Suivante</Button>
      <Button disabled={!document || zoom <= .5} aria-label="Réduire le zoom" onClick={() => setZoom(value => Math.max(.5, value - .25))}>−</Button>
      <Typography variant="body2">{Math.round(zoom * 100)} %</Typography>
      <Button disabled={!document || zoom >= 2} aria-label="Agrandir le zoom" onClick={() => setZoom(value => Math.min(2, value + .25))}>+</Button>
    </Box>
    <Box ref={container} sx={{ width: '100%', maxHeight: '65vh', overflow: 'auto', bgcolor: 'grey.200', p: 1 }}>
      {!document && !error && <CircularProgress size={24} sx={{ m: 3 }} />}
      <canvas ref={canvas} role="img" aria-label={`Page ${page} du document PDF`} style={{ display: document ? 'block' : 'none', margin: '0 auto' }} />
    </Box>
  </Box>
}
