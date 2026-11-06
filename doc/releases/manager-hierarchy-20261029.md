# Manager reports in Mon équipe — 2026-10-29

Published release: `rh-hierarchy-20261029T165827Z`.

Production read-only diagnosis confirmed that Robin's account has only `Responsable`, and its direct manager is William's separate `william.tang@ilci.fr` account. Both are active and neither linked personnel record is archived. William's separate Super Admin account remains unchanged.

Two independent filters caused the omission: the backend directory reused an attendance-edit rule requiring the target to have `Employe`; the frontend team table again filtered out every person without `Employe`.

Fix: the default authorized account directory also includes a manager's direct reports regardless of their business function. `team/teamPeople.ts` includes employees and managers in the team table while retaining inactive/deleted/archived filtering. The table label now says `Personnes gérées`. Attendance edits and account/role management continue to use their existing authorization rules; this change does not grant managerial peers edit access or include unrelated managers solely because they share a department.

Regression coverage includes a pure Responsable direct report, an unrelated Responsable in the same department, unchanged attendance/account edit denials, legacy roles and inactive/archived/deleted filtering. No production hierarchy, roles, contracts or personal data were changed to fix the display.

Verification: backend 497 tests, frontend 59 tests, lint and production build passed. Candidate checks confirmed Robin appears in the directory for William's separate Responsable account. After switching, 390 API checks passed across React and retained legacy APIs; 11 public assets including the TeamPage chunk matched release hashes. All 40 original production tables were preserved, all services are active and API startup logs have no traceback. No real mail or signature requests were sent by deployment checks.

Protected release backup: `/var/backups/presence-react/rh-hierarchy-20261029T165827Z`. Exact published files are identified by the release hash manifest, based on the current working tree. Coordinated and UI rollback scripts retain all business data.
