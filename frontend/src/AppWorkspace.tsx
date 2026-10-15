import { useCapabilities } from './capabilities'
import { useCallback, useState } from 'react'
import { Alert, Avatar, Box, Button, Chip, Drawer, IconButton, Typography } from '@mui/material'
import CalendarMonthRounded from '@mui/icons-material/CalendarMonthRounded'
import EventAvailableOutlined from '@mui/icons-material/EventAvailableOutlined'
import GroupsOutlined from '@mui/icons-material/GroupsOutlined'
import FolderOutlined from '@mui/icons-material/FolderOutlined'
import AdminPanelSettingsOutlined from '@mui/icons-material/AdminPanelSettingsOutlined'
import LogoutRounded from '@mui/icons-material/LogoutRounded'
import MenuRounded from '@mui/icons-material/MenuRounded'
import { ThemeToggle } from './components/ThemeToggle'
import { initials } from './leaveFormat'
import type { Session } from './types'
import { useDraftGuard, useWorkspaceGuard } from './draftGuard'
import { DraftGuardProvider } from './components/DraftGuardProvider'
import { recoverableModule } from './components/ModuleBoundary'
const AttendancePage = recoverableModule<Parameters<typeof import('./attendance/AttendancePage').AttendancePage>[0]>(() => import('./attendance/AttendancePage').then(module => ({ default: module.AttendancePage })))
const LeaveWorkspace = recoverableModule<Parameters<typeof import('./LeaveWorkspace').LeaveWorkspace>[0]>(() => import('./LeaveWorkspace').then(module => ({ default: module.LeaveWorkspace })))
const DocumentsPage = recoverableModule<Parameters<typeof import('./documents/DocumentsPage').DocumentsPage>[0]>(() => import('./documents/DocumentsPage').then(module => ({ default: module.DocumentsPage })))
const TeamPage = recoverableModule<Parameters<typeof import('./team/TeamPage').TeamPage>[0]>(() => import('./team/TeamPage').then(module => ({ default: module.TeamPage })))
const AdminPage = recoverableModule<Parameters<typeof import('./admin/AdminPage').AdminPage>[0]>(() => import('./admin/AdminPage').then(module => ({ default: module.AdminPage })))
const DashboardPage = recoverableModule<Parameters<typeof import('./dashboard/DashboardPage').DashboardPage>[0]>(() => import('./dashboard/DashboardPage').then(module => ({ default: module.DashboardPage })))
const ProfilePage = recoverableModule<Parameters<typeof import('./profile/ProfilePage').ProfilePage>[0]>(() => import('./profile/ProfilePage').then(module => ({ default: module.ProfilePage })))
import DashboardOutlined from '@mui/icons-material/DashboardOutlined'
import SettingsOutlined from '@mui/icons-material/SettingsOutlined'

type Page = 'attendance' | 'leave' | 'documents' | 'team' | 'admin' | 'dashboard' | 'profile'

export function AppWorkspace({ token, session, logout, onSession }: { token: string; session: Session; logout: () => void; onSession: (token: string) => void }) {
  return <DraftGuardProvider><Workspace token={token} session={session} logout={logout} onSession={onSession} /></DraftGuardProvider>
}

function Workspace({ token, session, logout, onSession }: { token: string; session: Session; logout: () => void; onSession: (token: string) => void }) {
  const capabilities = useCapabilities()
  const local = capabilities.data?.environment === 'local-snapshot'
  const attendanceAllowed = session.role_tags.some(role => ['Admin', 'Responsable'].includes(role))
  const [page, setPage] = useState<Page>(attendanceAllowed ? 'attendance' : 'leave')
  const [visited, setVisited] = useState<Partial<Record<Page, boolean>>>({ attendance: attendanceAllowed, leave: !attendanceAllowed })
  const [attendanceState, setAttendanceState] = useState({ dirty: false, busy: false })
  const onAttendanceState = useCallback((state: { dirty: boolean; busy: boolean }) => setAttendanceState(state), [])
  const guard = useWorkspaceGuard()!
  useDraftGuard(attendanceState.dirty, attendanceState.busy, 'Fiches de présence')
  const [drawer, setDrawer] = useState(false)
  function go(next: Page) { setPage(next); setVisited(v => ({ ...v, [next]: true })); setDrawer(false) }
  function navigate(next: Page) {
    if (next === page) { setDrawer(false); return }
    guard.request(() => go(next))
  }
  function requestLogout() { guard.request(logout) }
  const nav = <>
    <Box className="brand"><Box className="brand-mark">P</Box><Box><Typography sx={{ fontWeight: 750, fontSize: 21 }}>Presence</Typography><Typography sx={{ color: 'var(--sidebar-muted)', fontSize: 11, letterSpacing: 2 }}>ILCI · ESPACE RH</Typography></Box></Box>
    <Typography className="nav-heading">GESTION DES PRÉSENCES</Typography>
    {attendanceAllowed && <Button className={`nav-item ${page === 'dashboard' ? 'active' : ''}`} startIcon={<DashboardOutlined />} onClick={() => navigate('dashboard')}>Tableau de bord</Button>}
    {attendanceAllowed && <Button className={`nav-item ${page === 'attendance' ? 'active' : ''}`} startIcon={<CalendarMonthRounded />} onClick={() => navigate('attendance')}>Création de Fiches</Button>}
    <Button className={`nav-item ${page === 'leave' ? 'active' : ''}`} startIcon={<EventAvailableOutlined />} onClick={() => navigate('leave')}>Congés payés</Button>
    <Button className={`nav-item ${page === 'documents' ? 'active' : ''}`} startIcon={<FolderOutlined />} onClick={() => navigate('documents')}>Documents</Button>
    <Button className={`nav-item ${page === 'team' ? 'active' : ''}`} startIcon={<GroupsOutlined />} onClick={() => navigate('team')}>Mon équipe</Button>
    {session.role_tags.includes('Admin') && <Button className={`nav-item ${page === 'admin' ? 'active' : ''}`} startIcon={<AdminPanelSettingsOutlined />} onClick={() => navigate('admin')}>Administration</Button>}
    <Button className={`nav-item ${page === 'profile' ? 'active' : ''}`} startIcon={<SettingsOutlined />} onClick={() => navigate('profile')}>Paramètres</Button>
    <Box className="sidebar-note"><Typography sx={{ fontWeight: 650, fontSize: 13 }}>{local ? 'Version locale de test' : 'Espace Presence'}</Typography><Typography sx={{ color: 'var(--sidebar-muted)', fontSize: 12, mt: 1, lineHeight: 1.7 }}>{local ? 'Vos modifications restent dans la copie locale de test.' : 'Présences, documents et équipe dans votre espace.'}</Typography></Box>
    <Box className="sidebar-user"><Avatar sx={{ width: 36, height: 36, bgcolor: 'var(--nav-active)', color: 'var(--sidebar-text)', fontSize: 13 }}>{initials(session.name)}</Avatar><Box sx={{ flex: 1, minWidth: 0 }}><Typography noWrap sx={{ fontSize: 13, fontWeight: 650 }}>{session.name}</Typography><Typography sx={{ fontSize: 11, color: 'var(--sidebar-muted)' }}>{session.role_tags.join(' · ')}</Typography></Box><IconButton aria-label="Se déconnecter" disabled={guard.busy} onClick={requestLogout} sx={{ color: 'var(--sidebar-muted)' }}><LogoutRounded fontSize="small" /></IconButton></Box>
  </>
  return <Box className="app-shell">
    <Box component="aside" className="sidebar">{nav}</Box>
    <Drawer open={drawer} onClose={() => setDrawer(false)} slotProps={{ paper: { sx: { bgcolor: 'var(--sidebar)', color: 'var(--sidebar-text)', width: 260, p: 2, display: 'flex' } } }}>{nav}</Drawer>
    <Box className="main-column">
      <Box component="header" className="topbar"><IconButton className="mobile-menu" aria-label="Ouvrir le menu" onClick={() => setDrawer(true)}><MenuRounded /></IconButton>
        <Typography variant="body2" color="text.secondary">Espace RH <span className="breadcrumb-separator">/</span><strong>{({ attendance: 'Création de Fiches', documents: 'Documents', leave: 'Congés payés', team: 'Mon équipe', admin: 'Administration', dashboard: 'Tableau de bord', profile: 'Paramètres' })[page]}</strong></Typography>
        <Box sx={{ ml: 'auto', display: 'flex', alignItems: 'center', gap: 1 }}><ThemeToggle /><Chip className="preview-chip" label={local ? 'Copie locale' : capabilities.data ? 'Serveur' : 'Configuration à vérifier'} variant="outlined" size="small" /></Box>
      </Box>

      {guard.dirtyLabels.length > 0 && <Alert severity="info" role="status" sx={{ mx: { xs: 2, md: 5 }, mt: 2 }}>
        <Typography variant="body2" sx={{ fontWeight: 650 }}>Modifications non enregistrées</Typography>
        <Typography variant="body2">{guard.dirtyLabels.join(' · ')}. Enregistrez ou appliquez vos modifications avant de quitter.</Typography>
      </Alert>}

      {attendanceAllowed && <Box sx={{ display: page === 'attendance' ? 'block' : 'none' }}><AttendancePage token={token} session={session} onState={onAttendanceState} /></Box>}
      {visited.leave && <Box sx={{ display: page === 'leave' ? 'block' : 'none' }}><LeaveWorkspace embedded token={token} session={session} logout={requestLogout} /></Box>}
      {visited.documents && <Box sx={{ display: page === 'documents' ? 'block' : 'none' }}><DocumentsPage token={token} session={session} /></Box>}
      {visited.team && <Box sx={{ display: page === 'team' ? 'block' : 'none' }}><TeamPage token={token} session={session} onAttendance={() => navigate('attendance')} /></Box>}
      {visited.admin && <Box sx={{ display: page === 'admin' ? 'block' : 'none' }}><AdminPage token={token} session={session} /></Box>}
      {visited.dashboard && <Box sx={{ display: page === 'dashboard' ? 'block' : 'none' }}><DashboardPage token={token} session={session} /></Box>}
      {visited.profile && <Box sx={{ display: page === 'profile' ? 'block' : 'none' }}><ProfilePage token={token} session={session} onSession={onSession} /></Box>}

    </Box>

  </Box>
}
