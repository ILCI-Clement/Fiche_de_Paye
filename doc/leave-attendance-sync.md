# Paid leave and attendance synchronization

Local implementation; not yet deployed.

- Submission does not alter attendance. Applicable half-days show a pending hint.
- Approval updates scheduled working half-days, preserving rest days and French public holidays. Half-day boundaries and cross-month requests are supported.
- Missing month sheets are created in the approver's workspace with the existing contract/month binding and planning defaults. Missing contract coverage blocks the entire operation and the UI offers the attendance/contract completion entry. No contract dates are guessed.
- Existing matching sheets in other workspaces are updated in the same transaction; unrelated employees and month snapshots remain unchanged. Later-created copies inherit already approved leave on save.
- Existing paid leave is linked without alteration. Sickness, absence and other conflicting statuses require a fresh preview and explicit checkbox confirmation. The server rechecks the fingerprint under database locks.
- Approving cancellation restores only the half-days changed by that request. Unrelated later edits remain untouched. If the affected cell or schedule was changed, cancellation is blocked for manual review. Overlapping approved requests keep their leave until the last applicable request is cancelled.
- Production signatures protect their source month as soon as a request is tracked, including signatures in progress. A signed/archived sheet already displaying the correct paid leave is linked without changing its contents. Any mismatch blocks automatic approval/cancellation. PDFs and signature records are never regenerated or modified by this feature.
- Regular attendance saves also reject content changes to protected sheets. Archival metadata remains separate. Source locks serialize signature generation/provider acceptance with leave/calendar transactions; source content is reread after locking and compared with the preview fingerprint.
- Rejection, withdrawal and rejected cancellations do not modify attendance. No balance deduction or notification emails are introduced. Historical approvals are not automatically rewritten on startup.

## Storage and endpoints

`leave_attendance_cells` stores the request ID, owning workspace, sheet ID, date/half-day, original calendar state, applied state and whether the request actually changed it. This is a restoration/audit ledger, not a second employee profile. Rows survive request archival and follow workspace account renames. No cascade deletion is introduced.

`GET /leave-requests/{id}/attendance-preview?operation=approve|cancel` requires current review authority. The existing decision/cancellation-decision payload optionally carries `attendance_confirmation` (SHA-256 fingerprint). Blocked or stale operations return 409 with no partial writes. Existing authorization remains enforced on the server. The approval and its calendar/ledger/history changes use one transaction; audit failure rolls everything back.

## Validation

- `test_leave_attendance.py`: isolated synthetic SQLite HTTP transitions, cross-month half-days, multiple workspaces, missing contracts, protected correct/mismatched sheets, explicit conflicts, stale previews, cancellation, overlap, later-created copies and audit rollback.
- `test_responsable_workflow_mysql.py`: opt-in local MySQL connection-owned temporary tables, real auth/roles, approval/create/calendar/cancel, pending hints, revision conflicts and manual signed-source protection. Permanent business tables are not mutated.
- `test_attendance_signatures.py`: source reread after locking, changed-source rejection and provider call before releasing the lock; provider/generation are mocked, with no real mail or signature request.
- Browser verification uses one clearly labelled local TEST request, reads the approval preview, then withdraws it without approving or editing attendance.

Deploy only with the additive schema migration, the new Python modules and rebuilt frontend together. Back up per `backend/README.md`; do not run local test data or runtime configuration against production.

Final local verification: 538 backend tests passed with opt-in MySQL temporary-table integration enabled; 63 frontend tests, lint and production build passed. Local API startup succeeded. Two clearly labelled browser TEST requests were previewed and withdrawn without approval, attendance edits, emails or signature sends. Screenshot: ignored .local-test/leave-sync-preview.png. Production was not changed.

Published 2026-10-30 (Europe/Paris): `rh-file-batch-20261030T131557Z`; see [release verification](releases/file-batch-layout-20261030.md). Earlier local-only notes describe the pre-release verification stage.
