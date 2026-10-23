import { useCapabilities } from './capabilities'
import { brand } from './brand'
import { canManageTeam, displayRoles, isAdmin } from './rolePolicy'
import { useCallback, useLayoutEffect, useRef, useState } from 'react'
import type { ReactNode } from 'react'
import { Avatar, Box, Button, Chip, Drawer, IconButton, Tooltip, Typography } from '@mui/material'
import CalendarMonthRounded from '@mui/icons-material/CalendarMonthRounded'
import EventAvailableOutlined from '@mui/icons-material/EventAvailableOutlined'
import GroupsOutlined from '@mui/icons-material/GroupsOutlined'
import FolderOutlined from '@mui/icons-material/FolderOutlined'
import AdminPanelSettingsOutlined from '@mui/icons-material/AdminPanelSettingsOutlined'
import BadgeOutlined from '@mui/icons-material/BadgeOutlined'
import LogoutRounded from '@mui/icons-material/LogoutRounded'
import MenuRounded from '@mui/icons-material/MenuRounded'
import HelpOutlineRounded from '@mui/icons-material/HelpOutlineRounded'
import { ThemeToggle } from './components/ThemeToggle'
import { initials } from './leaveFormat'
import type { Session } from './types'
import { useDraftGuard, useWorkspaceGuard } from './draftGuard'
import { DraftGuardProvider } from './components/DraftGuardProvider'
import { DraftStatus } from './components/DraftStatus'
import { recoverableModule } from './components/ModuleBoundary'
import { WorkspaceToolbarContext } from './workspaceToolbar'
import { workspacePages } from './workspacePages'
import type { WorkspacePage } from './workspacePages'
const AttendancePage = recoverableModule<Parameters<typeof import('./attendance/AttendancePage').AttendancePage>[0]>(() => import('./attendance/AttendancePage').then(module => ({ default: module.AttendancePage })))
const LeaveWorkspace = recoverableModule<Parameters<typeof import('./LeaveWorkspace').LeaveWorkspace>[0]>(() => import('./LeaveWorkspace').then(module => ({ default: module.LeaveWorkspace })))
const DocumentsPage = recoverableModule<Parameters<typeof import('./documents/DocumentsPage').DocumentsPage>[0]>(() => import('./documents/DocumentsPage').then(module => ({ default: module.DocumentsPage })))
const TeamPage = recoverableModule<Parameters<typeof import('./team/TeamPage').TeamPage>[0]>(() => import('./team/TeamPage').then(module => ({ default: module.TeamPage })))
const AdminPage = recoverableModule<Parameters<typeof import('./admin/AdminPage').AdminPage>[0]>(() => import('./admin/AdminPage').then(module => ({ default: module.AdminPage })))
const DashboardPage = recoverableModule<Parameters<typeof import('./dashboard/DashboardPage').DashboardPage>[0]>(() => import('./dashboard/DashboardPage').then(module => ({ default: module.DashboardPage })))
const ProfilePage = recoverableModule<Parameters<typeof import('./profile/ProfilePage').ProfilePage>[0]>(() => import('./profile/ProfilePage').then(module => ({ default: module.ProfilePage })))
const PersonnelPage = recoverableModule<Parameters<typeof import('./personnel/PersonnelPage').PersonnelPage>[0]>(() => import('./personnel/PersonnelPage').then(module => ({ default: module.PersonnelPage })))
import DashboardOutlined from '@mui/icons-material/DashboardOutlined'
import SettingsOutlined from '@mui/icons-material/SettingsOutlined'

type Page = WorkspacePage

function WorkspacePageSlot({ active, target, inlineTarget, children }: { active: boolean; target: HTMLElement | null; inlineTarget: HTMLElement | null; children: ReactNode }) {
  return <WorkspaceToolbarContext value={{ active, target, inlineTarget }}><Box hidden={!active} sx={{ display: active ? 'block' : 'none', minWidth: 0 }}>{children}</Box></WorkspaceToolbarContext>
}

export function AppWorkspace({ token, session, logout, onSession }: { token: string; session: Session; logout: () => void; onSession: (token: string) => void }) {
  return <DraftGuardProvider><Workspace token={token} session={session} logout={logout} onSession={onSession} /></DraftGuardProvider>
}

function Workspace({ token, session, logout, onSession }: { token: string; session: Session; logout: () => void; onSession: (token: string) => void }) {
  const capabilities = useCapabilities()
  const local = capabilities.data?.environment === 'local-snapshot'
  const attendanceAllowed = canManageTeam(session)
  const personnelAllowed = isAdmin(session)
  const [page, setPage] = useState<Page>(attendanceAllowed ? 'attendance' : 'leave')
  const adminPage = personnelAllowed && (page === 'personnel' || page === 'admin')
  const [visited, setVisited] = useState<Partial<Record<Page, boolean>>>({ attendance: attendanceAllowed, leave: !attendanceAllowed })
  const [attendanceState, setAttendanceState] = useState({ dirty: false, busy: false })
  const onAttendanceState = useCallback((state: { dirty: boolean; busy: boolean }) => setAttendanceState(state), [])
  const guard = useWorkspaceGuard()!
  useDraftGuard(attendanceState.dirty, attendanceState.busy, 'Fiches de présence')
  const [drawer, setDrawer] = useState(false)
  const mainColumn = useRef<HTMLDivElement>(null)
  const header = useRef<HTMLElement>(null)
  const [toolbarTarget, setToolbarTarget] = useState<HTMLDivElement | null>(null)
  const [inlineToolbarTarget, setInlineToolbarTarget] = useState<HTMLDivElement | null>(null)
  useLayoutEffect(() => {
    const element = header.current, column = mainColumn.current
    if (!element || !column) return
    const measure = () => column.style.setProperty('--workspace-toolbar-height', `${Math.ceil(element.getBoundingClientRect().height)}px`)
    measure()
    // Observe the header only: changing the body height cannot create a resize loop.
    if (typeof ResizeObserver !== 'undefined') {
      const observer = new ResizeObserver(measure)
      observer.observe(element)
      return () => observer.disconnect()
    }
    window.addEventListener('resize', measure)
    return () => window.removeEventListener('resize', measure)
  }, [])
  function go(next: Page) { setPage(next); setVisited(v => ({ ...v, [next]: true })); setDrawer(false) }
  function navigate(next: Page) {
    if (next === page) { setDrawer(false); return }
    guard.request(() => go(next))
  }
  function requestLogout() { guard.request(logout) }
  const nav = <>
    <Box className="brand"><Box className="brand-mark">{brand.mark}</Box><Box><Typography sx={{ fontWeight: 750, fontSize: 21 }}>{brand.name}</Typography><Typography sx={{ color: 'var(--sidebar-muted)', fontSize: 11, letterSpacing: 2 }}>{brand.organization} · PORTAIL RH</Typography></Box></Box>
    <Typography className="nav-heading">RESSOURCES HUMAINES</Typography>
    {attendanceAllowed && <Button className={`nav-item ${page === 'dashboard' ? 'active' : ''}`} startIcon={<DashboardOutlined />} onClick={() => navigate('dashboard')}>Tableau de bord</Button>}
    {attendanceAllowed && <Button className={`nav-item ${page === 'attendance' ? 'active' : ''}`} startIcon={<CalendarMonthRounded />} onClick={() => navigate('attendance')}>Création de Fiches</Button>}
    <Button className={`nav-item ${page === 'leave' ? 'active' : ''}`} startIcon={<EventAvailableOutlined />} onClick={() => navigate('leave')}>Congés payés</Button>
    <Button className={`nav-item ${page === 'documents' ? 'active' : ''}`} startIcon={<FolderOutlined />} onClick={() => navigate('documents')}>Documents</Button>
    <Button className={`nav-item ${page === 'team' ? 'active' : ''}`} startIcon={<GroupsOutlined />} onClick={() => navigate('team')}>Mon équipe</Button>
    <Button className={`nav-item ${page === 'profile' ? 'active' : ''}`} startIcon={<SettingsOutlined />} onClick={() => navigate('profile')}>Paramètres</Button>
    {personnelAllowed && <Box component="section" aria-label="Espace Admin" className={`admin-nav-group${adminPage ? ' admin-nav-group--active' : ''}`}>
      <Typography component="h2" className="admin-nav-heading"><AdminPanelSettingsOutlined aria-hidden="true" sx={{ fontSize: 16 }} />Espace Admin</Typography>
      <Button className={`nav-item ${page === 'personnel' ? 'active' : ''}`} aria-current={page === 'personnel' ? 'page' : undefined} startIcon={<BadgeOutlined />} onClick={() => navigate('personnel')}>Registre du personnel</Button>
      <Button className={`nav-item ${page === 'admin' ? 'active' : ''}`} aria-current={page === 'admin' ? 'page' : undefined} startIcon={<AdminPanelSettingsOutlined />} onClick={() => navigate('admin')}>Administration</Button>
    </Box>}
    <Box className="sidebar-note"><Typography sx={{ fontWeight: 650, fontSize: 13 }}>{local ? 'Version locale de test' : brand.title}</Typography><Typography sx={{ color: 'var(--sidebar-muted)', fontSize: 12, mt: 1, lineHeight: 1.7 }}>{local ? 'Vos modifications restent dans la copie locale de test.' : 'Personnel, contrats, documents et congés dans votre espace.'}</Typography></Box>
    <Box className="sidebar-user"><Avatar sx={{ width: 36, height: 36, bgcolor: 'var(--nav-active)', color: 'var(--sidebar-text)', fontSize: 13 }}>{initials(session.name)}</Avatar><Box sx={{ flex: 1, minWidth: 0 }}><Typography noWrap sx={{ fontSize: 13, fontWeight: 650 }}>{session.name}</Typography><Typography noWrap title={displayRoles(session.role_tags)} sx={{ fontSize: 11, color: 'var(--sidebar-muted)' }}>{displayRoles(session.role_tags)}</Typography></Box><IconButton aria-label="Se déconnecter" disabled={guard.busy} onClick={requestLogout} sx={{ color: 'var(--sidebar-muted)' }}><LogoutRounded fontSize="small" /></IconButton></Box>
  </>
  return <Box className="app-shell">
    <Box component="aside" className="sidebar">{nav}</Box>
    <Drawer open={drawer} onClose={() => setDrawer(false)} slotProps={{ paper: { sx: { bgcolor: 'var(--sidebar)', color: 'var(--sidebar-text)', width: 260, p: 2, display: 'flex' } } }}>{nav}</Drawer>
    <Box ref={mainColumn} className="main-column">
      <Box component="header" ref={header} className="topbar topbar--with-tools"><Box className="topbar-primary"><IconButton className="mobile-menu" aria-label="Ouvrir le menu" onClick={() => setDrawer(true)}><MenuRounded /></IconButton>
        <Tooltip key={page} title={workspacePages[page].description} describeChild enterDelay={650} enterNextDelay={650} arrow>
          <Typography component="h1" className="topbar-title" tabIndex={0} sx={{ m: 0, minWidth: 0, display: 'inline-flex', alignItems: 'center', gap: .75, fontSize: 17, fontWeight: 600, lineHeight: 1.4, cursor: 'help', overflowWrap: 'anywhere', '&:focus-visible': { outline: '2px solid', outlineColor: 'primary.main', outlineOffset: 4, borderRadius: 1 } }}>{workspacePages[page].title}<HelpOutlineRounded aria-hidden="true" sx={{ fontSize: 15, flexShrink: 0, color: 'text.secondary' }} /></Typography>
        </Tooltip>
        {adminPage && <Chip className="admin-area-label" icon={<AdminPanelSettingsOutlined />} label="Espace Admin" size="small" variant="outlined" color="primary" />}
        <Box ref={setInlineToolbarTarget} className="topbar-inline-tools" />
        <Box sx={{ ml: 'auto', display: 'flex', alignItems: 'center', gap: 1 }}><ThemeToggle /><Chip className="preview-chip" label={local ? 'Copie locale' : capabilities.data ? 'Serveur' : 'Configuration à vérifier'} variant="outlined" size="small" /></Box>
        </Box>
        <Box ref={setToolbarTarget} className="topbar-tools" />
      </Box>

      {page !== 'attendance' && <DraftStatus />}

      {attendanceAllowed && <WorkspacePageSlot active={page === 'attendance'} target={toolbarTarget} inlineTarget={inlineToolbarTarget}><AttendancePage token={token} session={session} onState={onAttendanceState} /></WorkspacePageSlot>}
      {visited.leave && <WorkspacePageSlot active={page === 'leave'} target={toolbarTarget} inlineTarget={inlineToolbarTarget}><LeaveWorkspace token={token} session={session} /></WorkspacePageSlot>}
      {visited.documents && <WorkspacePageSlot active={page === 'documents'} target={toolbarTarget} inlineTarget={inlineToolbarTarget}><DocumentsPage token={token} session={session} /></WorkspacePageSlot>}
      {visited.team && <WorkspacePageSlot active={page === 'team'} target={toolbarTarget} inlineTarget={inlineToolbarTarget}><TeamPage token={token} session={session} onAttendance={() => navigate('attendance')} /></WorkspacePageSlot>}
      {visited.admin && <WorkspacePageSlot active={page === 'admin'} target={toolbarTarget} inlineTarget={inlineToolbarTarget}><AdminPage token={token} session={session} /></WorkspacePageSlot>}
      {visited.dashboard && <WorkspacePageSlot active={page === 'dashboard'} target={toolbarTarget} inlineTarget={inlineToolbarTarget}><DashboardPage token={token} session={session} /></WorkspacePageSlot>}
      {visited.profile && <WorkspacePageSlot active={page === 'profile'} target={toolbarTarget} inlineTarget={inlineToolbarTarget}><ProfilePage token={token} session={session} onSession={onSession} /></WorkspacePageSlot>}
      {personnelAllowed && visited.personnel && <WorkspacePageSlot active={page === 'personnel'} target={toolbarTarget} inlineTarget={inlineToolbarTarget}><PersonnelPage token={token} session={session} /></WorkspacePageSlot>}

    </Box>

  </Box>
}
