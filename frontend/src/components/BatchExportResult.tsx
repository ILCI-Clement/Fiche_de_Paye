import { Accordion, AccordionDetails, AccordionSummary, Alert, Box, Button, LinearProgress, Typography } from '@mui/material'
import ExpandMoreRounded from '@mui/icons-material/ExpandMoreRounded'
import { downloadFile } from '../download'

export interface ExportResult {
  phase: 'preparing' | 'ready' | 'error'; title: string; items: string[]; filename?: string; error?: string;
}

/** ZIP APIs are atomic: do not invent per-file successes when the batch fails. */
export function BatchExportResult({ result }: { result: ExportResult }) {
  return <Box role="status" sx={{ my: 2 }}>
    <Alert severity={result.phase === 'error' ? 'error' : 'info'}>
      <Typography variant="body2" sx={{ fontWeight: 650 }}>{result.title}</Typography>
      {result.phase === 'preparing' ? `Préparation de ${result.items.length} élément(s)… Le téléchargement n’est pas encore prêt.`
        : result.phase === 'ready' ? `${result.items.length} élément(s) préparé(s). Téléchargement déclenché : ${result.filename}. Vérifiez les téléchargements de votre navigateur.`
        : `Export interrompu : aucun fichier ni ZIP partiel n’a été téléchargé. ${result.error || ''} La sélection est conservée pour correction ou nouvel essai.`}
    </Alert>
    {result.phase === 'preparing' && <LinearProgress aria-label="Préparation de l’export" />}
    <Accordion disableGutters>
      <AccordionSummary expandIcon={<ExpandMoreRounded />}><Typography variant="body2">Sélection de cet export ({result.items.length})</Typography></AccordionSummary>
      <AccordionDetails>
        <Box component="ul" sx={{ my: 0, pl: 2.5, maxHeight: 220, overflow: 'auto', overflowWrap: 'anywhere' }}>{result.items.map((label, index) => <li key={index}>{label}</li>)}</Box>
        {result.phase !== 'preparing' && <Button sx={{ mt: 1 }} onClick={() => downloadFile({ filename: 'bilan-export.json', blob: new Blob([JSON.stringify(result, null, 2)], { type: 'application/json' }) })}>Télécharger le bilan</Button>}
      </AccordionDetails>
    </Accordion>
  </Box>
}
