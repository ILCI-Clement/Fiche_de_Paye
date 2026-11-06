# Leave notification emails — 2026-11-04

Published: `rh-mailnotify-20261104T162306Z`. The sole API/current symlink and Nginx static root now use this release, replacing API `rh-reports-20261104T140753Z` and UI `rh-report-align-20261104T142250Z`.

New leave requests and cancellation requests email the current eligible direct manager. Approval/rejection of either emails the employee, including result, dates, decision timestamp/comment and a login-protected details link. Withdrawal sends no email. No eligible manager means the ordinary Admin pending task remains, without popup or email broadcast. No historical request/event backfill.

The additive InnoDB `leave_notification_outbox` uses audit-event IDs to prevent duplicate jobs. Enqueue shares the business transaction; network delivery happens afterward. Known pre-send errors retry up to five attempts; uncertain SMTP outcomes or abandoned sending claims are never automatically repeated. Operators should inspect failed/uncertain jobs before any manual resend. No applicant note or private personnel number is included in mail. Details access and all decisions still require normal authenticated authorization.

Validation: 607 backend tests (including real MySQL temporary-table workflows), 76 frontend tests, lint and production build passed. Earlier explicitly authorized one-off synthetic mail was received by the user, who confirmed its local login/details link. During deployment the candidate worker was explicitly disabled; no real test mail/signing request was sent. Candidate/current each passed 367 main checks plus existing security/mailbox/storage checks. Ten public resource hashes matched. Runtime confirms public HTTPS links, active SMTP configuration, enabled notification worker and transactional empty queue. Original 42 business tables preserved; only the outbox table was added, bringing total to 43. Roles and real requests were not changed for verification.

All four background workers are enabled only on the sole production API. Retired services remain masked, old ports closed. Protected full backup: `/var/backups/presence-react/rh-mailnotify-20261104T162306Z`. Guarded `rollback.py` restores the preceding API/UI without restoring business data or deleting the additive queue. A rollback to the previous API suspends notification processing; retain queue state and inspect before re-enabling to avoid stale or duplicate mail.

Evidence: candidate-verified.json, published-checks.json, public-check.json, mail-runtime-check.json and verification-summary.json in the protected backup.
