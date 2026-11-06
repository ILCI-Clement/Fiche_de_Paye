# Signed PDF deletion

Implemented locally on 2026-10-29; published 2026-10-30 (Paris). See [release verification](releases/access-export-20261030.md).

`DELETE /signed-fiches/{archive_id}` accepts `{"confirmation":"SUPPRIMER"}` (old ID confirmations remain compatible). The server locks users, re-reads the actor's current account/session and requires Super Admin; it then locks the archive row. Only `status=ready` (signed PDF already archived) is eligible. Pending, uncertain, partially signed, completed-but-downloading and terminal unsigned requests are rejected. This avoids racing the callback/archive worker and does not cancel external signature requests.

The safe archive path is validated before mutation. A minimal `signed_fiche` entry in the existing independent permanent-deletion ledger and a `permanently_deleted` signature event are committed with removal of the document row. Original signature events remain for audit. Database or audit failures leave the PDF intact. After commit the local archived PDF is unlinked; a filesystem error returns `cleanup_pending=true`, leaves the row inaccessible and displays a maintenance warning. No recursive deletion is used. Deleted references cannot be restored by later callbacks, downloaded or included in monthly ZIP exports. ClawShow originals, attendance sheets, accounts and other archives are unchanged.

`GET /signed-fiches` returns `can_delete` from the freshly checked role. `SignedFichesPanel.tsx` shows a small delete action only for ready rows when that flag is true; the shared `PermanentDeleteDialog` requires the short word `SUPPRIMER`. Successful deletion invalidates signed lists, monthly profiles and dashboard work-center caches. The server remains authoritative when UI permissions become stale.

Verification: 30 signed-archive HTTP tests on synthetic SQLite/PDF fixtures; complete backend suite 508 passed, frontend suite 61 passed, lint and production build passed. The real MySQL TEMPORARY-table integration also exercises deletion, row locking, file removal and the independent audit. Local browser inspection opened and cancelled the confirmation for a simulation PDF; no existing local or production file was deleted through the UI, no emails/signature requests were sent.

Manual folders and batch actions: [file batch actions](file-batch-actions.md).
