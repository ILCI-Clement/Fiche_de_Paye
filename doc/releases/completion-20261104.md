# Completion navigation and concise help — 2026-11-04

Release: `rh-completion-20261104T104651Z`.

Readiness actions retain exact account/employee IDs and carry a completion issue code. Manager issues expand and focus organization settings; linked personnel issues open the correct personnel editor; missing entry dates open the employment editor; unlinked accounts open association search. Missing contracts open the existing month-specific completion dialog. Focus/scroll is form-scoped, waits for visible controls, briefly highlights the target and cancels on unmount. No automatic writes, social-security-number reveal, new permissions or business validation changes. Manually returning from an email action to generic completion resets focus to the current missing field.

Routine repeated prose/status chips were removed or moved into hover descriptions on account, personnel, association, employment/history and settings views. Required labels, validation errors, protected actions, local mail-disabled notice and destructive confirmations remain.

Verification before publishing: 69 frontend tests, lint and TypeScript/production build passed. Disposable read-only browser fixtures verified manager, name, email, birth date, employment entry date, contract start and unlinked association focus; no live records edited, no emails or signing requests sent. Fixture files/tab removed.

All packaged backend/dependency files are byte-identical to verified production `rh-storage-20261104T093711Z`; only built frontend assets change. Local private data/test fixtures are excluded.

Protected backup: `/var/backups/presence-react/rh-completion-20261104T104651Z`, including database, private files, signature SQLite, source/configuration and key material. Guarded `rollback.py` restores the sole API/UI release to `rh-storage-20261104T093711Z`; it never restores old services or business data.

Publication verification will be appended after cutover.

## Published verification

Published: candidate and current API each passed 367 main checks, plus mailbox, signature/contact and storage authorization checks. All 41 original business tables retained their data. Ten public assets matched the release hashes; anonymous private routes remain protected. Current API/Nginx are active, all three existing workers remain enabled, retired services remain masked with old ports closed. Production roles and data were not changed by this release. No test emails or signing requests sent. Protected records: candidate-verified.json, published-checks.json, public-check.json, verification-summary.json.
