# Retired standalone API compatibility

2026-10-30: production uses only `backend.rebuild_server:app`. This source snapshot preserves the modules before removing their standalone flat-import branches, plus the retired flat-import test and the old revision-test scaffolding. It contains no runtime configuration, database, keys or user data.

`manifest.json` records archive and per-file SHA-256 hashes. It is reference material, not a runnable production release. Do not restore these files over the current project: that would also revert newer mailbox identity and permission safeguards. For historical reference, extract into an independent directory only.

Current package imports, local testing and historical data compatibility remain supported. Shared `backend/main.py` is still an active module; only its unused functions and obsolete flat-import branches were removed.
