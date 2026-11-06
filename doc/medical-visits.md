# Visites medicales - local development

Independent menu alongside Documents. Records actual occupational visit dates and
visit attestations, not diagnoses or clinical reports. Date required, PDF optional
(up to 10 MiB). Employees upload their own evidence; Admin can register on behalf
of visible personnel. Missing personnel association displays a contact-admin state.
No appointments, expiry reminders or approval workflow. Not deployed.

## Current access rules

- Employee: own currently linked personnel dossier, current and previous versions.
- Responsable: own dossier plus current managed employees, using the existing
  `main.user_can_manage` callback shared with team/attendance permissions. Read-only
  for subordinates, including previous versions and PDF evidence. Changing managers
  immediately changes access to all old visits; no historical-manager restriction.
- Admin/Super Admin: all visible personnel dossiers, including people without a login.
  Existing Super Admin self-only catalog visibility is enforced.
- Default catalogs hide disabled/deleted logins and archived personnel. Historical
  access to an archived personnel dossier remains with Admin and the owner, not the
  former team manager. Individual archived visits have a separate Archives filter.
- Only owner/Admin may correct dates/PDFs, while dossier and visit are unarchived.
  Admin can archive/restore visits; only Super Admin can permanently delete after
  archiving, version confirmation and typing SUPPRIMER. Deletion removes the record,
  versions, detailed visit events and encrypted PDFs; a minimal independent deletion
  ledger retains actor, resource ID and counts only. Filesystem cleanup failure is
  returned explicitly as cleanup_pending; no accessible evidence remains in the DB.

`medical_access.py` reuses live team relations for read access, without extending
write access. Creation choices use an own-only catalog for non-Admin, so a manager
cannot accidentally submit for the first subordinate returned by a list.

## Corrections and persistence

`medical_visits.py` owns creation, scoped lists/history, attachments, file reads and
existing ZIP export. `medical_history.py` owns version list/file routes, correction,
archive/restore and guarded deletion. Registered in local/rebuild entries.

Additive InnoDB tables: medical_visits, medical_visit_events and
medical_visit_versions. Startup adds is_archived to older local medical tables;
existing visits/files are preserved. First attachment, correction, archive and
permanent deletion check the latest version under the same user/visit locks.
A correction saves the previous date, file reference and version, with actor/time,
then writes the new version in the same transaction. Old encrypted files remain
available through authorized version routes. Failed writes/audits roll back and
remove only their new file. Existing register deletion refuses referenced dossiers;
foreign keys prevent orphaned history.

Every route reloads current roles, activity and session version. URLs contain no
credentials. Responses use the shared private/no-store policy. PDF reads, version
reads and exports are audited without storing PDF contents in event logs. Normal
Documents routes cannot access these files. Actual PDF parsing rejects unreadable
or password-protected files.

## Encryption and environment

`medical_crypto.py` uses AES-256-GCM with a separate key ring. Associated data binds
purpose/version, employee ID, random filename and key ID.

- PRESENCE_MEDICAL_DIR: private encrypted directory, default
  /var/lib/presence-app/medical-visits.
- PRESENCE_MEDICAL_KEYS_FILE: private JSON ring in the existing personnel key-ring
  format. Keep old keys on rotation. No production key is generated automatically.
  Missing keys reject PDF reads/writes; date-only corrections still work.

Local entry uses independent keys/files in .local-test/rebuild. Before deployment,
provision the private production key/directory and include all three tables/files in
backups, protecting key backups separately. Browser downloads contain plaintext PDFs.

## Period filters and existing export

Year/optional month use actual visit date, not upload time. Summaries show the latest
visit within the selected period; history has month headings. Registration choices
are independent of the period. Existing POST /medical-visits/export uses matching
search, period, archive and current access rules across all pages, with ZIP layout
YYYY/MM/Employee name_ID/YYYY-MM-DD_visitID_filename.pdf. Maximum 100 PDF / 100 MiB.
Missing/corrupt files fail the whole export without partial audit or partial ZIP.
Excel links and merged PDF are implemented locally; details below.

## Verification

`test_medical_visits.py`: synthetic HTTP cases including scope transfer with the
production team policy, subordinate read-only access, own registration choices,
versioned correction, stale write rejection, archive/restore, Super-only deletion,
rollback/file cleanup, authenticated encryption and bounded all-page ZIP exports.
`test_medical_visits_mysql.py`: opt-in loopback-only disposable database checks
idempotent schema, FK retention, real correction/version/PDF reads, manager transfer
and guarded permanent deletion. Run scripts/run_tests.py --mysql with these modules.
UI reuses the existing draft guard, French date field, PDF preview and permanent
confirmation dialog. No real employee records are used for destructive tests.


## ZIP export (local only)

Exports contain only original PDFs under YYYY/MM/Name_employeeID_date_visitID.pdf. No index PDF, Excel summary or file links. Stable IDs prevent duplicate-name collisions; filenames use shared sanitization. Original PDF bytes remain unchanged.

The toolbar exports the current year/month/search/archive scope across all pages. Each employee row downloads that person's entire history, all years including archived visits. History checkboxes export selected certificates. The backend requires employee_id for all_history and checks current owner/team/Admin permissions for every selected ID. Maximum 100 PDFs / 100 MiB. Missing files or generation errors fail atomically; export audits roll back and no partial ZIP is returned. Downloads remain private/no-store. Not deployed.


Published 2026-11-05 as `rh-medical-20261105T161729Z`. Earlier local-only status is superseded; see [release verification](releases/medical-visits-20261105.md). No local fixtures were imported.
