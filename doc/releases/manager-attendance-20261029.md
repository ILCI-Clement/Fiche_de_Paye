# Attendance for a directly reporting manager — 2026-10-29

Release: `rh-manager-attendance-20261029T171527Z`.

The earlier team-directory fix did not change attendance authorization, so Robin still did not appear in William's attendance selector. This release explicitly permits an active, non-Admin Responsable who reports directly to another Responsable to be managed for their own attendance. Same-department peers are not included solely because both are managers.

`role_policy.direct_business_report` is shared by `main.user_can_manage` and missing-contract setup in `contract_history.py`. The existing attendance catalog, monthly creation, calendar editing, save validation and signature contact checks reuse that scope. Work-center readiness includes these direct manager reports. Admin/Super Admin delegation and account editing remain guarded separately.

`main.require_workspace_access` prevents a superior manager from reading or replacing a subordinate manager's entire saved workspace through the legacy get-config/save-config routes. That workspace may contain other people's records. The superior creates the subordinate's personal attendance in the superior's own workspace.

Regression: 497 backend tests passed. The synthetic MySQL workflow covers selection, missing-contract response, permitted contract setup, creation, calendar application, save/reload, peer-manager denial, privileged-account denial, role-edit denial and subordinate workspace isolation. No real Robin contract dates or attendance data are created or modified by deployment verification.

Protected backup and guarded rollback: `/var/backups/presence-react/rh-manager-attendance-20261029T171527Z`. The exact release hash manifest identifies the working-tree package; local runtime, synthetic records and keys are excluded.

Published successfully. Candidate checks explicitly verified Robin in William's attendance workspace account catalog and denied access to Robin's whole legacy workspace. Post-switch: 393 API checks passed, 10 public resources matched release hashes, all services active and no API startup traceback. Concurrent genuine register and attendance edits triggered the data checkpoint guard; old backups were retained and a final protected backup was captured with API writers briefly stopped before switching. All 40 tables matched that checkpoint at publication. Roles, passwords and real contract/attendance values were not edited by this fix or its verification.
