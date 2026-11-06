# Direct report name/department alignment — 2026-11-04

Published frontend: `rh-report-align-20261104T142250Z`. The Nginx static root points to this release's `frontend/dist`. The sole API/current symlink remains `rh-reports-20261104T140753Z`; no service restart or API/source/data change.

ReportsEditor receives existing department metadata from UserEditor, renders names in equally sized outlined chips, and aligns department badges in the following column. Narrow screens use two flexible columns; long names retain full hover titles. No assignment or permission behavior changes.

Lint and build passed. Publication verified 250 on-disk resource hashes, nine public assets including AdminPage, three protected private routes and 20 authenticated read-only account/profile checks. API PID and all 455 previous release files remained unchanged. No emails, signature sends or business writes.

Protected backup: `/var/backups/presence-react/rh-report-align-20261104T142250Z`. `rollback-ui.sh` restores the previous Nginx static root only; it checks that the expected UI is active and leaves API/data untouched. Backup includes prior static assets, frontend source, database, private files and source/configuration.
