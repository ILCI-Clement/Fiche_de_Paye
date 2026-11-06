# React UI reference archive

`CreateAccount.tsx.txt` preserves the removed standalone Administration account-creation dialog, including its manual-password option. It has no active imports and is excluded from the frontend source/build. The source was moved intact on 2026-10-29; its relative imports are historical references, not a runnable component here.

Active creation now lives in `frontend/src/personnel/PersonnelEditor.tsx`: employee registration with optional linked account, or external account only. Both reuse `RegistrationAccountFields` and the shared API client. Existing backend account authorization, hashing and post-commit email flow remain active.
