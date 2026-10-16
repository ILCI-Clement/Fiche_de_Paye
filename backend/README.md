# API Presence

React remake released 2026-10-15 as `react-20261015T101334Z`. Public UI: `https://app.presence.217.65.145.103.nip.io/`. Service `presence-react-api` runs the full-repository entry `backend.rebuild_server:app` on loopback 8015. Legacy services remain active for rollback. See `../doc/rebuild-release-readiness.md` for backup paths and the guarded rollback command.

This directory contains the version-controlled FastAPI service used by the Presence interface.

## Deployment

## Paid leave module (React service)

The React interface in `../frontend` registers `/leave/session` and `/leave-requests`
through `leave_requests.py`. Deploy `leave_rules.py`, `leave_store.py`, and
`leave_requests.py` with the new `main.py`; startup creates the two leave tables.
The entry point also updates request ownership on account rename and detaches it
on account deletion, preserving immutable operation history. Back up both source
and MariaDB before restarting, and verify authorization and transactions first.

The retired fictitious-account preview was removed on 2026-10-16. Use
`local_rebuild.py` with the isolated authorized data snapshot for local work.
The production React entry remains `rebuild_server.py`, reusing `main.py`.

## Production environment

The production service reads its credentials from `/etc/presence-app/presence-app.env`, via the systemd drop-in for `presence-app.service`. Do not commit that file or any password, SMTP credential, database credential, or signing secret.

## Signature électronique

`esign.py` converts attendance sheets to PDF and requests an employee signature
followed by the responsible manager through ClawShow eSign. `signature_flow.py`
resolves and validates the two people. Both signatures and their placement were
verified in a completed sandbox employee PDF; the user subsequently confirmed the
live employee-to-manager signing flow. The complete modular version was deployed
on 2026-10-14. The intern end-to-end signing flow remains to be verified.
The production environment must provide `CLAWSHOW_ESIGN_API_KEY`;
it must only exist in `/etc/presence-app/presence-app.env`. LibreOffice is required
on the API server for the Excel/Word-to-PDF conversion. Generated PDFs are kept in
`PRESENCE_ESIGN_DOCUMENTS_DIR` (default: `/var/lib/presence-app/esign-documents`)
behind an unguessable URL for 24 hours, so ClawShow can retrieve the document.
`PRESENCE_ESIGN_PUBLIC_URL` must be the public FastAPI origin (not the Streamlit
interface URL), for example `https://presence.example.org`.

## Documents

Transport receipts are stored privately in `PRESENCE_TRANSPORT_RECEIPTS_DIR`
(default: `/var/lib/presence-app/transport-receipts`). The API accepts one PDF,
JPG or PNG up to 10 MB at a time. The database stores only file metadata,
assignment, document type, bound month/year and archive state; downloads remain protected by the authenticated API.
At startup, the `transport_receipts` table is created if it does not exist.
`document_numbers.py` adds a nullable unique `document_number` and a
`document_daily_sequences` counter table. It fills missing references in upload
order (`created_at`, then `id`), using the timestamp converted to Europe/Paris.
New uploads reserve `YYYYMMDD-0001` references in the same transaction as the
metadata insert, using a per-day InnoDB row lock. Archive/month edits retain the
reference; deletion never decrements the counter. At least four sequence digits
are displayed, extending naturally past 9999. Single downloads and ZIP entries
include the reference; private storage filenames and original names stay intact.
Deploy `document_numbers.py` with `documents.py` for both active API entries.
Rollback may retain the additive column, unique index and counters, avoiding any
loss of issued references; do not restore a database dump over later uploads.

Each authenticated employee or manager may upload only for their own username.
For non-admins, `/transport-receipts/assignees` returns only that username and the
upload route rejects a different target with HTTP 403 before writing files or
database rows. Admin retains upload-on-behalf permission. Managers retain team
read access, without additional archive, period-correction or deletion rights.
Existing document ownership is unchanged.

`documents.py` owns the existing `/transport-receipts` routes. The name stays for
URL and client compatibility. Frontend ZIP exports use these individual authenticated
downloads, enforce a 100 MB limit, and do not return incomplete archives.

Before replacing the deployed API, take both a source backup and a MariaDB dump. At startup, the API performs only additive schema migrations: it adds optional organization columns to `users` and creates the group tables if they do not already exist. It does not delete existing user accounts or `Presence` records.

The expected command is:

```text
uvicorn main:app --host 127.0.0.1 --port 8001
```

The deployed service uses `/var/www/presence-app/main.py`. For the flat server layout,
deploy `main.py`, `documents.py`, `esign.py`, `contract_history.py`,
`attendance_permissions.py`, `signature_flow.py` and `signature_placement.py` together in that directory.
The team-directory update also requires `team_directory.py` with the updated
`main.py`. It adds authenticated read-only organization and own-department directory
routes; existing personnel, contract and Documents authorization remains unchanged.
Local package imports (`backend.main`) and flat imports (`main`) are both supported.
Route registration injects the original authentication and database dependencies;
submodules must not import `main.py` back.

Signature placement requires `pypdf` from `requirements.txt`. The current ClawShow
signing-page renderer maps field coordinates through a 595 by 842 reference page.
`signature_placement.py` finds the labels in the final PDF and sends normalized
explicit coordinates, without mixing anchors and explicit page/x/y fields. This
avoids the observed landscape-anchor shift. Recheck this mapping if the provider
changes its field coordinate convention. Rotated/cropped or ambiguous templates
are rejected rather than signed at a guessed location.

Frontend deployment must include the changed pages and their common helpers:
`api_client.py`, `fiche_exports.py`, `fiche_serialization.py`, `planning.py`,
`calendar_view.py`, `document_ui.py` and `document_downloads.py`, alongside existing
export/template and period modules. Compare with production before overwriting files;
the server may contain changes absent locally. Replacing only a page or `main.py`
can leave its new imports missing.

The attendance batch-selection update also requires `fiche_batch_ui.py` together
with `pages/Fiches.py`. Batch download and signing are separate actions using the
same explicit selection. Signing retains the employee's registered direct manager.
Completion webhook processing and automatic signed-PDF archival are not implemented.

The module map is in [../doc/modules.md](../doc/modules.md). Current refactoring and
verification status is recorded in [../doc/refactoring.md](../doc/refactoring.md).

## Runtime configuration

Configure values in the protected server environment; the table lists names only.
Do not copy that environment into the repository or frontend secrets.

| Variables | Purpose |
| --- | --- |
| `PRESENCE_API_TOKEN` | Required session signing secret |
| `PRESENCE_DB_HOST`, `PRESENCE_DB_USER`, `PRESENCE_DB_PASSWORD`, `PRESENCE_DB_DATABASE` | Required MariaDB connection |
| `PRESENCE_SMTP_SERVER`, `PRESENCE_SMTP_PORT`, `PRESENCE_SMTP_USER`, `PRESENCE_SMTP_PASSWORD` | Required mail configuration |
| `PRESENCE_APP_URL` | Required frontend origin for account emails |
| `PRESENCE_TRANSPORT_RECEIPTS_DIR` | Private long-term Documents storage; default given above |
| `PRESENCE_ESIGN_DOCUMENTS_DIR` | Temporary signing PDFs; default given above |
| `PRESENCE_ESIGN_PUBLIC_URL` | Public HTTPS API origin from which ClawShow fetches PDFs |
| `CLAWSHOW_ESIGN_API_KEY`, `CLAWSHOW_ESIGN_NAMESPACE` | Provider key; namespace defaults to `ilci` |
| `PRESENCE_ESIGN_ORGANIZATION_NAME`, `PRESENCE_ESIGN_ORGANIZATION_FULL_NAME` | Mail institution identity; defaults to ILCI and its full name |
| `PRESENCE_ESIGN_REPRESENTATIVE_NAME`, `PRESENCE_ESIGN_REPRESENTATIVE_TITLE` | Mail closing name and position; defaults to resolved manager and `Responsable` |

The frontend reads `URL_PRESENCE` from its own `.streamlit/secrets.toml`. A local
UI can connect to the production API, so actions there affect real server data.
To run only the local UI, use the project virtual environment:

```powershell
.\.venv\Scripts\python.exe -m streamlit run online_test.py --server.address 127.0.0.1 --server.port 8502 --server.headless true
```

## Release and rollback

The 2026-10-14 modular release backup is
`/var/backups/presence-interface/online-release-20261014T091212Z`. It contains a
file manifest, the prior deployed source files and a MariaDB dump. Runtime keys
were preserved in place and are not part of that source release payload.

For each future release, compare source differences, back up every replaced file
and the database, validate flat backend imports, and deploy all required modules.
Keep existing file ownership and modes. Restart both services, verify their active
state, public API routes and Streamlit health, and inspect startup errors.

For a source rollback, stop the services, restore the manifest's existing files
from that release backup, remove only files marked newly added by that release,
and restart and verify the services. Leave uploaded Documents and server secrets
intact. Additive schema changes do not require automatic database restoration;
restoring a database dump can discard later writes and needs a separate decision.

Health checks establish startup and reachability. They do not establish successful
login, uploads, export contents or email delivery. Known behavior needing correction,
including destructive account deletion and private session caches, is tracked in
[the maintenance plan](../doc/code-audit.md).


## React server entry (released 2026-10-15)

`backend.rebuild_server:app` is the server entry for the full React migration.
Run from the complete repository root; retain root-level generators and templates.
It reuses `main.app`, authentication, current permissions and checked saves, then
registers the React attendance and signature adapters. Do not deploy the isolated
`local_rebuild.py` entry or its runtime/snapshot files.

The entry uses the existing server environment and private document paths. Set
`PRESENCE_SIGNATURE_ATTEMPTS_DB` to a writable persistent SQLite file (default
`/var/lib/presence-app/signature-attempts.sqlite3`) and back it up with the service.
Readiness validates the selected signing key prefix and public HTTPS PDF origin;
it does not prove external PDF accessibility or email delivery. No provider call
is made by this readiness check. Signing receipt storage is not final PDF archival.

`ui_capabilities.py` registers public `GET /ui-capabilities`, returning only
boolean feature availability and environment type. Local startup explicitly
reports account creation and mail as disabled and blocks the corresponding routes.
The capability endpoint does not replace authorization on business endpoints.

Uninstalled systemd/Nginx examples are in `../deploy/`. They use a separate API
port and preserve the old service for rollback. Review the actual server paths,
service user, proxy prefixes, SSL host, permissions and environment before use.
The new entry still inherits the original startup migrations and reminder worker;
starting it against production is a release action, requiring backup and approval.
See `../doc/rebuild-release-readiness.md` for acceptance and remaining boundaries.

The new service disables its reminder worker using `PRESENCE_CONTRACT_REMINDER_WORKER=0` while the retained legacy API owns the worker. If retiring the legacy API later, enable this worker on the new service. Manual reminders remain authorized normally. Source rollback switches the UI and stops the new service; it does not restore the database dump or remove uploaded files.
