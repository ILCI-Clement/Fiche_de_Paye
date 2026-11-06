# UI cleanup — 2026-10-29

This pass removes redundant active code without changing storage or backend authorization:

- Moved the unreferenced standalone `admin/CreateAccount.tsx` intact to `archive/react/CreateAccount.tsx.txt`. Unified registration and external account creation remain the active UI. The archive preserves the manual-password form for reference.
- Combined employee and external submission completion in `PersonnelEditor`: one action lock, one success cleanup, one account-created result. External submission still sends only names, email and organization to `/create-user`; no register/contract/NIR payload reaches that endpoint. Employee validation and transactional registration remain unchanged.
- Removed the obsolete `visited` map from `AppWorkspace`. `WorkspacePageSlot` still mounts only the active module, preserving immediate draft disposal. Admin and dashboard slots now explicitly follow the same permission predicates as their menu entries.

Active TypeScript source: 120 files / 5,369 lines → 119 files / 5,289 lines (80 fewer active lines). This measures active source only; the archived form is intentionally retained. A relative-import reachability check found no further unreferenced TypeScript modules from main.tsx.

Frontend lint, TypeScript, 54 tests and production build passed. Backend regression uses isolated fixtures and local connection-owned MySQL TEMPORARY tables; no production changes or real mails/signatures. This cleanup is local only.
