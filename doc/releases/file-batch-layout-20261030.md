# File batching, leave synchronization and shared header release

Published 2026-10-30 (Europe/Paris), immutable UTC release ID `rh-file-batch-20261030T131557Z`.

## Delivered

- Pending leave withdrawal and approved-leave cancellation requests, centered confirmations and clickable summary filters.
- Transactional approved-leave synchronization by half-day, restoration with overlapping-request safeguards, missing-sheet creation/contract checks and protection of signed attendance sources. See [business rules](../leave-attendance-sync.md).
- Annual signed-PDF ZIP export, grouped into month folders; pagination moved to the Utilisateurs panel footer.
- Documents / Signé batch archive, restore and Super Admin-only permanent deletion. Signé has Normal / Archives folders; short `SUPPRIMER` confirmation replaces the internal ID in the UI. See [batch rules](../file-batch-actions.md).
- Theme/environment controls moved to the sidebar footer; redundant sidebar prose removed. Shared desktop header aligns title and page actions with a 72px minimum height; crowded or narrow layouts wrap without horizontal page overflow.

## Verification

- 555 backend tests passed, including isolated real MySQL temporary-table workflows; 63 frontend tests, lint and production build passed.
- No server source differences against the previous release manifest `rh-access-export-20261029T232232Z`.
- Protected backup created before candidate startup. Candidate: 361 checks across 12 existing accounts; all original 40 tables retained with unchanged existing data.
- Additive schema changes include leave cancellation state, the half-day synchronization ledger and signed-PDF manual-folder state. Final checkpoint includes 41 tables; the original baseline remains in the protected backup.
- Published APIs on 8015 and 8001: 618 checks passed. Included live permissions, refused batch mutations against nonexistent references, archive list access, signed history boundaries, PDF/Excel export permissions and manager attendance visibility. No real file deletions, exports of employee PII, emails or signature requests were performed by release tests.
- Ten public resources matched their manifest hashes. Nginx, both APIs and retained interface active. Current API invocations have no traceback.
- No local database/simulation files imported; no production passwords or roles changed. Production NIR protection remains enabled on the React API.

## Recovery

Current symlink and Nginx root: `/var/www/presence-react/releases/rh-file-batch-20261030T131557Z` (UI under `frontend/dist`). Retained flat API modules updated consistently. No dependency installation was required.

Protected backup: `/var/backups/presence-react/rh-file-batch-20261030T131557Z`. Includes database checkpoints, source/configuration, private files, signature SQLite, snapshot keyring copy, baseline manifests and verification records. API writers were briefly stopped for the final checkpoint; no concurrent business differences were found.

`rollback.sh` coordinates the previous code and UI, leaving additive schema and all subsequent business writes intact. `rollback-ui.sh` restores only the previous UI. Do not restore the database over later production writes.
