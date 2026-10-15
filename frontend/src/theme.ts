import { createTheme } from '@mui/material/styles'

export function buildTheme(mode: 'light' | 'dark') {
  const dark = mode === 'dark'
  return createTheme({
  palette: { mode,
    primary: { main: dark ? '#7aa7f8' : '#285fbd', contrastText: dark ? '#101827' : '#ffffff' },
    secondary: { main: dark ? '#b9a9df' : '#6f5aa2' },
    background: { default: dark ? '#0e1117' : '#f2f3f5', paper: dark ? '#262730' : '#fafafa' },
    text: { primary: dark ? '#fafafa' : '#31333f', secondary: dark ? '#b8bcc7' : '#666b79' },
    divider: dark ? '#3d404b' : '#d9dce4',
  },
  typography: { fontFamily: '"Segoe UI", -apple-system, BlinkMacSystemFont, sans-serif',
    h4: { fontSize: '1.9rem', fontWeight: 750, letterSpacing: '-0.8px' },
    h6: { fontWeight: 700 }, button: { textTransform: 'none', fontWeight: 650 } },
  shape: { borderRadius: 8 },
  components: {
    MuiAccordion: { defaultProps: { defaultExpanded: false }, styleOverrides: {
      root: { border: `1px solid ${dark ? '#3d404b' : '#d9dce4'}`, borderRadius: '8px !important', '&::before': { display: 'none' }, '&.Mui-expanded': { margin: '12px 0' } },
    } },
    MuiAccordionSummary: { styleOverrides: { root: { minHeight: 52, paddingInline: 18, '&.Mui-expanded': { minHeight: 52 } }, content: { margin: '12px 0', '&.Mui-expanded': { margin: '12px 0' } } } },
    MuiAccordionDetails: { styleOverrides: { root: { padding: '8px 18px 20px' } } },
    MuiAlert: { styleOverrides: { root: ({ ownerState }) => ownerState.severity === 'success' ? {
      backgroundColor: dark ? '#343740' : '#e9ebef', color: dark ? '#fafafa' : '#31333f',
      border: `1px solid ${dark ? '#555966' : '#c7ccd5'}`,
      '& .MuiAlert-icon': { color: 'inherit' },
    } : {} } },
    MuiCssBaseline: { styleOverrides: {
      ':root': {
        '--canvas': dark ? '#0e1117' : '#f2f3f5', '--surface': dark ? '#262730' : '#fafafa',
        '--surface-soft': dark ? '#30323d' : '#e9ebf0', '--surface-hover': dark ? '#30323d' : '#f0f1f4',
        '--border': dark ? '#3d404b' : '#d9dce4', '--muted': dark ? '#b8bcc7' : '#666b79',
        '--sidebar': dark ? '#262730' : '#e9ebf0', '--accent': dark ? '#7aa7f8' : '#285fbd',
        '--accent-contrast': dark ? '#101827' : '#ffffff',
        '--sidebar-text': dark ? '#fafafa' : '#31333f', '--sidebar-muted': dark ? '#b8bcc7' : '#666b79',
        '--nav-active': dark ? '#2b3b55' : '#dde8fa', '--nav-hover': dark ? '#353743' : '#e0e3ea',
      },
    } },
    MuiButton: { defaultProps: { disableElevation: true }, styleOverrides: { root: { minHeight: 44, borderRadius: 8, paddingInline: 16 } } },
    MuiOutlinedInput: { styleOverrides: { root: { minHeight: 44 }, input: { fontSize: 14 } } },
    MuiTab: { styleOverrides: { root: { minHeight: 48, paddingInline: 18, fontWeight: 650 } } },
    MuiTableCell: { styleOverrides: { root: { padding: '12px 16px', fontSize: 13 }, head: { color: dark ? '#b8bcc7' : '#666b79', backgroundColor: dark ? '#20222b' : '#eef0f4', fontWeight: 650, whiteSpace: 'nowrap' } } },
    MuiTableRow: { styleOverrides: { root: { '&:last-child td': { borderBottom: 0 }, '&:hover td': { backgroundColor: dark ? '#ffffff04' : '#00000003' } } } },
    MuiTextField: { defaultProps: { size: 'small', variant: 'outlined' } },
    MuiPaper: { defaultProps: { elevation: 0 }, styleOverrides: { root: { backgroundImage: 'none' } } },
    MuiChip: { styleOverrides: { root: { fontWeight: 650, fontSize: 12 } } },
    MuiDialog: { styleOverrides: { paper: { borderRadius: 18 } } },
  },
  })
}
