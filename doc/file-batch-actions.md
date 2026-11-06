# File batch actions

Local implementation, 2026-10-30; not deployed by this change.

- `POST /transport-receipts/batch`: selected Documents, including all filtered pages; maximum 100 IDs. Existing ownership rules apply to Standard users, Responsable remains read-only, Admin can archive/restore, and only Super Admin can permanently delete already archived receipts.
- `POST /signed-fiches/batch`: selected ready PDFs on the current 20-item page. Admin can archive/restore; only Super Admin can permanently delete. Normal / Archives is a manual folder distinction, independent from successful PDF retrieval (`status=ready`, `archived_at`). Moving folders never rewrites the signed PDF, changes source attendance protections or removes historical audit events. Complete annual/monthly ZIPs continue to include both folders.
- Both payloads contain `ids`, `action` (`archive`, `restore`, `delete`) and strict boolean `expected_archived`; deletion also requires `confirmation: "SUPPRIMER"`. Existing single-delete endpoints accept this short word and retain old ID confirmations for client compatibility.
- Current actor/session and rows are re-read under transaction locks. Every selected row is checked before mutation. Missing rows, changed state, invalid paths or audit failure roll back the complete database batch. Physical deletion happens after commit. `cleanup_pending=true` explicitly reports a filesystem cleanup failure; the deleted row remains inaccessible.
- `backend/file_batch.py` shares payload rules and postcommit cleanup. Frontend `components/FileBatchDialog.tsx` freezes the submitted selection, lists files and handles confirmation/errors; `fileBatch.ts` shares labels and API types. Successful operations clear selection and invalidate document/signature, profile and task caches.

Verification: 555 backend tests passed, including MySQL TEMPORARY-table migration and operations; 63 frontend tests passed, lint and production build passed. Thirteen dedicated synthetic batch tests exercise ownership, live role revocation, rollback, path checks, payload bounds, cleanup warnings and short confirmations. Browser inspection opens and cancels confirmation dialogs; no real files, external signatures or emails are changed.

Published 2026-10-30 (Europe/Paris): `rh-file-batch-20261030T131557Z`; see [release verification](releases/file-batch-layout-20261030.md). Earlier local-only notes describe the pre-release verification stage.
