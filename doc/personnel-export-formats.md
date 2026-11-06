# User-register download formats

Implemented locally on 2026-10-29; published 2026-10-30 (Paris). See [release verification](releases/access-export-20261030.md).

- Admin: current filtered register and historical snapshots download as PDF.
- Super Admin: PDF or template-based Excel, selected in the download dialog.
- Other roles: neither register export endpoint nor snapshot download is available.

`POST /personnel-register/export.xlsx` now requires the freshly re-read Super Admin role. `POST /personnel-register/export.pdf` requires Admin or Super Admin. Both reuse the same complete filtered query, 10,000-row limit, employment-period rows, NIR confirmation and transactional per-person export audit. PDF uses the same in-memory template cell values as Excel; neither plaintext file is persisted. `pdf_exported` events store format/count metadata, never numbers or business values.

`backend/personnel_pdf.py` converts template XLSX cells A:R, starting at row 8, to a static landscape A3 PDF with wrapped text, repeated column headers, page numbers and confidential footer. Dates and all 18 template columns are retained; sex is displayed as H/F. A PDF reduces direct spreadsheet editing but is not an unalterable or cryptographically certified record. Existing application edit permissions and audits are unchanged. ReportLab 4.x is declared in `backend/requirements.txt` and must be installed in the target environment before deployment.

Historical encrypted XLSX snapshot blobs and digests remain untouched. Snapshot downloads accept `format=pdf|xlsx`, default PDF; the server rejects Admin requests for XLSX before decrypting content. PDF conversion uses only the verified historical blob and original capture timestamp, not current employee data. The full/redacted NIR variant choice, sensitive confirmation and download audit remain. `can_export_excel` is returned from the live snapshot-list permission check.

Frontend entry: PersonnelToolbar now says Exporter; PersonnelExport receives Super Admin capability from the session and shows PDF/Excel only for that role. PersonnelSnapshots uses the server flag for historical format choices. Backend checks remain authoritative if the session UI is stale.

Validation: 513 backend tests, including local MySQL TEMPORARY-table snapshot and permission checks; 63 frontend tests, lint and production build. Additional export tests cover all filtered pages, fresh role checks, missing confirmations, audit failure rollback and PDF generation failure. Synthetic 44-row PDF was rendered and visually reviewed across pages for table layout, French accents, repeated headers and page numbering. No real employee export, email or signature submission was performed.
