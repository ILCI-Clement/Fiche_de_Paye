# Access controls and user export release - 2026-10-30

Released during the Paris date 2026-10-30. Immutable UTC release ID: `rh-access-export-20261029T232232Z`.

## Changes

- Super Admin account catalog entries are visible only to the owner, including other Super Admin viewers; historical authors and business records remain intact.
- Business function selection (Employé/Responsable) is separate from system access. Admin/Super Admin checkboxes are aligned with the account name and require confirmation before changing the draft; server-side delegation and self-protection remain authoritative.
- Signed and archived PDFs have a Super Admin-only permanent-delete action, exact reference confirmation and independent deletion audit. Unfinished signature requests are ineligible. ClawShow originals are unchanged.
- The user-register menu/title is now Utilisateurs.
- Dashboard account navigation no longer falls back to the first person when an explicit target is unavailable. Manager-contact corrections identify the manager and label the action accordingly.
- Admin register exports and snapshot downloads are PDF only. Super Admin can choose PDF or Excel. Backend format guards prevent direct XLSX bypass. PDF output uses the same 18 template columns, all matching pages/periods, repeated headers and page numbers. Encrypted historical XLSX blobs are unchanged; PDF is generated from the verified historical content in memory.

## Validation

- 513 backend regression tests passed, including connection-scoped synthetic MySQL temporary-table workflows.
- 63 frontend tests, lint and production build passed.
- No source conflicts against the previous deployed release `rh-manager-attendance-20261029T171527Z`.
- Candidate service on loopback 18015: 321 API checks across 12 existing accounts; all 40 production tables unchanged.
- Server PDF renderer generated and parsed a synthetic document without reading real employee export data.
- Published APIs on 8015 and retained 8001: 538 checks passed. Checks included role scopes, Super Admin identity hiding, Excel-denied-for-Admin, PDF access, snapshot format restrictions, denied deletion and nonexistent-reference probes, signature history isolation and William/Robin attendance scope.
- Ten public resources matched the exact manifest hashes; both APIs, retained interface and Nginx active.
- All 40 existing production tables were preserved. No local snapshot/simulation data imported, roles/passwords changed, real signatures sent, emails sent or real employee PDFs exported/deleted by release tests.

## Deployment and recovery

Manifested working-tree release: 306 files. Current symlink and Nginx root point to `/var/www/presence-react/releases/rh-access-export-20261029T232232Z`. The retained flat API modules were updated consistently. ReportLab 4.5.1 was added to the existing server virtual environment; previous dependency inventory and install/check logs are in the protected backup. This dependency addition does not rewrite business data.

Protected backup: `/var/backups/presence-react/rh-access-export-20261029T232232Z`. Database dump, source/configuration, private files, signature SQLite, dedicated keyring backup, baseline hashes and candidate/publish verification records were saved. API writers were briefly quiesced for the final full checkpoint; no concurrent data differences were detected at that checkpoint.

Guarded `rollback.sh` coordinates the previous API/UI code without restoring the database; `rollback-ui.sh` restores only the previous UI while retaining current backend guards. Subsequent business writes must not be overwritten by a database restore. The additive ReportLab package may remain installed during a code rollback.
