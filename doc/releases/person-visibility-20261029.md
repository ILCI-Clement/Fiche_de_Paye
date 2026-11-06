# Daily personnel visibility — 2026-10-29

Release: `rh-visibility-20261029T162813Z`.

`backend/person_visibility.py` supplies a shared daily-catalog rule: inactive or login-deleted accounts, and accounts linked to an archived employee-register record, are hidden. The register is optional on older installations; unexpected database errors still propagate.

Applied to default `/list-users`, `/my-team`, `/organization-directory`, attendance account catalogs, document upload assignees, dashboard personnel and readiness lists. Admin-only historical account queries retain all records and expose `personnel_archived` for the new Administration `Archivés` tab. Disabling authentication, archiving a register entry and retaining historical records remain separate operations. Contracts, signed documents, historical attendance and admin detail endpoints are preserved.

Register mutations invalidate team, organization, upload assignee, work-center and attendance caches in addition to the existing register/account caches. New manager and employee assignment options omit archived people; the currently assigned historical manager stays readable in the organization editor.

Validation: 496 backend tests passed, including a real MySQL TEMPORARY-table archive/restore workflow across admin and manager catalogs, team, organization and attendance, plus inactive-account checks. Document assignee filtering tested separately. Frontend 57 tests, lint, TypeScript and production build passed. Candidate: 213 API checks over 12 accounts; original contents of 40 production tables unchanged before switch.

Protected backup and guarded rollback scripts: `/var/backups/presence-react/rh-visibility-20261029T162813Z`. Release contents are identified by the exact hash manifest; no local simulation data or credentials are uploaded.

An additional explicit user request authorizes promotion of William's uniquely matched company account (`william.tang@univ-ilci.fr`) to Super Admin after deployment verification. The maintenance action preserves business roles and groups, records a role-only backup and audit event, and invalidates old sessions. No password change is involved.

Published successfully: 368 post-switch API checks passed; 10 public assets matched the package and all services are active with no startup traceback. A concurrent legitimate production update triggered the pre-switch data guard; the deployment stopped before switching, retained the original backup, captured an additional full database backup and refreshed checkpoint, then repeated candidate verification. The resulting 40-table checkpoint was preserved through publication.

William promotion completed: exact account username/email `william.tang@univ-ilci.fr`; roles changed from Admin/Responsable to SuperAdmin/Admin/Responsable. Both APIs accept the new session version and reject the previous one. Super Admin snapshot-cycle capability verified. Role-only backup: `william-roles-before.json`; outcome: `william-promotion.json`; audit action: `authorized_super_admin_granted`, actor `maintenance:christophe`, under the protected release backup directory. Other account and personnel data were not edited by this promotion.
