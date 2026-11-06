import { ReportsEditor } from './ReportsEditor'
import { storeRememberedToken } from '../authSession'
import { personDisplayName } from '../personDisplayName'
import { displayRoles, isAdmin, isSuperAdmin } from '../rolePolicy'
import { useDraftGuard } from '../draftGuard'
import { useId, useRef, useState } from 'react'
import { completionSection } from '../completionTarget'
import { useCompletionFocus } from '../components/useCompletionFocus'
import type { WorkRequest } from '../workspaceNavigation'
import { Alert, Box, Button, Dialog, DialogActions, DialogContent, DialogTitle, Tab, Tabs, Typography } from '@mui/material'
import type { Group } from '../team/api'
import { OrganizationFields } from './OrganizationFields'
import { SystemAccessFields } from './SystemAccessFields'
import { organizationValidation } from './organizationValidation'
import { adminApi } from './api'
import type { AdminUser, OrganizationInput, RunAction } from './api'
import { ContractEditor } from './ContractEditor'
import { AccountAccess } from './AccountAccess'
import { AccountPersonnelSection } from '../profile/AccountPersonnelSection'
import { AccountReadiness } from '../dashboard/AccountReadiness'
import { useWorkCenter } from '../dashboard/useWorkCenter'
import { CompletionDot } from '../components/CompletionDot'
import { accountStatus } from './accountStatus'

export function UserEditor({ token, user, users, groups, actor, actorRoles, busy, run, formId, requestError = '', completion }: { completion?: WorkRequest; token: string; user: AdminUser; users: AdminUser[]; groups: Group[]; actor: string; actorRoles: string[]; busy: boolean; run: RunAction; formId: string; requestError?: string }) {
  const [draft, setDraft] = useState<OrganizationInput>({ role_tags: user.role_tags, employee_type: user.employee_type, group_ids: user.group_ids, managed_group_ids: user.managed_group_ids, manager_id: user.manager_id })
  const [baseline, setBaseline] = useState(() => JSON.stringify(draft))
  const confirmChange = useDraftGuard(JSON.stringify(draft) !== baseline, busy, `Compte · ${personDisplayName(user)}`)
  const organizationRoot = useRef<HTMLDivElement>(null)
  const section = completion?.completion ? completionSection(completion.completion) : undefined
  const [expanded, setExpanded] = useState(section || 'register')
  const tabsId = useId()
  const [accessActions, setAccessActions] = useState<HTMLDivElement | null>(null)
  const [appliedCompletion, setAppliedCompletion] = useState(completion?.sequence)
  if (completion && appliedCompletion !== completion.sequence) { setAppliedCompletion(completion.sequence); if (section) setExpanded(section) }
  const [roleConfirm, setRoleConfirm] = useState(false)
  const [organizationError, setOrganizationError] = useState('')
  const [today] = useState(() => new Date())
  const readiness = useWorkCenter(token, actor, today.getMonth() + 1, today.getFullYear())
  const issues = readiness.data?.people.find(person => person.username === user.username)?.issues ?? []
  const registerIssues = issues.filter(issue => issue.domain === 'register' || ['name', 'email', 'manager_contact'].includes(issue.code)).map(issue => issue.label)
  const organizationIssues = issues.filter(issue => issue.code === 'manager').map(issue => issue.label)
  const organizationIssue = organizationValidation(draft)
  if (organizationIssue && !organizationIssues.includes(organizationIssue)) organizationIssues.push(organizationIssue)
  const contractIssues = issues.filter(issue => issue.domain === 'attendance').map(issue => issue.label)
  useCompletionFocus(organizationRoot, expanded === 'organization' && (section === 'organization' || organizationIssues.length > 0) ? 'manager' : undefined, completion?.sequence)
  function submitOrganization() {
    if (readOnly) return
    const error = organizationValidation(draft)
    setOrganizationError(error)
    if (error) return
    if ([...draft.role_tags].sort().join() !== [...user.role_tags].sort().join()) setRoleConfirm(true)
    else void saveOrganization()
  }
  async function saveOrganization() {
    if (protectedTarget) return
    const validation = organizationValidation(draft)
    setOrganizationError(validation)
    if (validation) return
    let failure = ''
    const ok = await run(async () => {
      try {
        const result = await adminApi.organization(token, user.username, draft)
        if (user.username === actor && [...draft.role_tags].sort().join() !== [...user.role_tags].sort().join()) {
          // The server revoked all old credentials. Do not silently restore this session.
          storeRememberedToken()
          window.dispatchEvent(new Event('presence:expired'))
        }
        return result
      }
      catch (error) { failure = error instanceof Error ? error.message : 'Enregistrement impossible.'; throw error }
    }, 'Organisation enregistrée.')
    if (ok) { setBaseline(JSON.stringify(draft)); setRoleConfirm(false) }
    else setOrganizationError(failure || 'Enregistrement non confirmé. Actualisez puis réessayez.')
  }
  const superAdmin = isSuperAdmin({ role_tags: actorRoles })
  const protectedTarget = isSuperAdmin(user) && !superAdmin
  const readOnly = busy || protectedTarget
  const reports = users.filter(p => p.manager_id === user.username && accountStatus(p) === 'active')
  const systemOnly = isSuperAdmin({ role_tags: draft.role_tags })
  const activeTab = systemOnly && (expanded === 'register' || expanded === 'contract') ? 'organization' : expanded
  return <>
    <Box sx={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', flexWrap: 'wrap', gap: 1, mb: 1 }}>
      <Typography variant="h6" sx={{ minWidth: 0, overflowWrap: 'anywhere' }}>Fiche de {personDisplayName(user)}</Typography>
      <Box ref={setAccessActions} />
      <Box sx={{ ml: 'auto' }}><SystemAccessFields roles={draft.role_tags} onChange={roles => setDraft({ ...draft, role_tags: roles })}
        disabled={readOnly} allowSuperAdmin={superAdmin} preserveAdmin={user.username === actor && isAdmin(user) && !superAdmin} /></Box>
    </Box>
    <Typography color="text.secondary" sx={{ mb: 2, overflowWrap: 'anywhere' }}>{user.email}</Typography>
    <AccountReadiness token={token} viewer={actor} username={user.username} />
    {protectedTarget && <Alert severity="info" sx={{ mb: 2 }}>Seul un Super Admin peut modifier ce compte.</Alert>}
    {organizationError && !roleConfirm && <Alert severity="error" sx={{ mb: 2 }}>{organizationError}</Alert>}
    <Tabs value={activeTab} onChange={(_, value: string) => setExpanded(value)} aria-label="Rubriques du compte"
      variant="scrollable" scrollButtons="auto" sx={{ borderBottom: 1, borderColor: 'divider', mb: 2,
        '& .MuiTab-root': { minHeight: 44, minWidth: 0, px: 2, textTransform: 'none' } }}>
      {!systemOnly && <Tab value="register" id={`${tabsId}-register`} aria-controls={`${tabsId}-register-panel`}
        label={<Box component="span" sx={{ display: 'inline-flex', alignItems: 'center', gap: 1, '& [role="img"]': { mr: 0 } }}>Informations<CompletionDot issues={registerIssues} /></Box>} />}
      <Tab value="organization" id={`${tabsId}-organization`} aria-controls={`${tabsId}-organization-panel`}
        label={<Box component="span" sx={{ display: 'inline-flex', alignItems: 'center', gap: 1, '& [role="img"]': { mr: 0 } }}>Organisation<CompletionDot issues={organizationIssues} /></Box>} />
      {!systemOnly && <Tab value="contract" id={`${tabsId}-contract`} aria-controls={`${tabsId}-contract-panel`}
        label={<Box component="span" sx={{ display: 'inline-flex', alignItems: 'center', gap: 1, '& [role="img"]': { mr: 0 } }}>Contrats et historique<CompletionDot issues={contractIssues} /></Box>} />}
      <Tab value="reports" id={`${tabsId}-reports`} aria-controls={`${tabsId}-reports-panel`} label={`Personnes sous responsabilité (${reports.length})`} />
    </Tabs>
    <Box sx={{ minHeight: { sm: 320 } }}>
      {!systemOnly && <Box role="tabpanel" hidden={activeTab !== 'register'} id={`${tabsId}-register-panel`} aria-labelledby={`${tabsId}-register`}>
        <AccountPersonnelSection key={user.username} token={token} username={user.username} viewer={actor} disabled={busy} embedded
          completionIssues={registerIssues} completion={section === 'register' ? completion : undefined} />
      </Box>}
      <Box role="tabpanel" hidden={activeTab !== 'organization'} id={`${tabsId}-organization-panel`} aria-labelledby={`${tabsId}-organization`}>
        <Box ref={organizationRoot} component="form" id={formId} onSubmit={e => { e.preventDefault(); submitOrganization() }} sx={{ display: 'grid', gridTemplateColumns: { xs: '1fr', md: '1fr 1fr' }, gap: 2 }}>
          <OrganizationFields value={draft} onChange={setDraft} users={users} groups={groups} username={user.username} disabled={readOnly} hideSystemAccess />
        </Box>
        <Box sx={{ mt: 3, pt: 1, borderTop: 1, borderColor: 'divider' }}>
        <AccountAccess actionsContainer={accessActions} token={token} user={user} actor={actor} actorRoles={actorRoles} busy={busy} run={run} confirmChange={confirmChange} requestError={requestError} />
      </Box>
      </Box>
      {!systemOnly && <Box role="tabpanel" hidden={activeTab !== 'contract'} id={`${tabsId}-contract-panel`} aria-labelledby={`${tabsId}-contract`}>
        <Box data-completion-field="contract"><ContractEditor token={token} username={user.username} readOnly={protectedTarget} canDelete={superAdmin} busy={busy} run={run} requestError={requestError} /></Box>
      </Box>}
      <Box role="tabpanel" hidden={activeTab !== 'reports'} id={`${tabsId}-reports-panel`} aria-labelledby={`${tabsId}-reports`}>
        <ReportsEditor token={token} manager={user} users={users} groups={groups} disabled={protectedTarget} busy={busy} run={run} />
      </Box>

    </Box>
    <Dialog open={roleConfirm} onClose={() => { if (!busy) setRoleConfirm(false) }}><DialogTitle>Confirmer le changement de droits</DialogTitle><DialogContent>
      {organizationError && <Alert severity="error" sx={{ mb: 2 }}>{organizationError}</Alert>}
      <Alert severity="warning" sx={{ mb: 2 }}>Les accès de {personDisplayName(user)} seront modifiés.</Alert>
      {user.username === actor && <Alert severity="warning" sx={{ mb: 2 }}>Vos droits seront modifiés et vous serez déconnecté. Reconnectez-vous pour utiliser les nouveaux droits.</Alert>}
      <Typography>Actuellement : {displayRoles(user.role_tags)}</Typography>
      <Typography>Après modification : {displayRoles(draft.role_tags)}</Typography>
    </DialogContent><DialogActions><Button disabled={busy} onClick={() => setRoleConfirm(false)}>Annuler</Button><Button variant="contained" disabled={busy} onClick={() => void saveOrganization()}>Confirmer les nouveaux droits</Button></DialogActions></Dialog>
  </>
}
