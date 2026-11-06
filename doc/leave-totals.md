# Approved paid leave totals (local only)

Profile uses the existing year selector and authorization (self, current direct manager, Admin). The annual total and monthly values count approved request half-days; .5 day per recorded half-day. Dates already passed and future approved dates both count. This is an approved-leave total, not earned entitlement, remaining balance or an hours conversion.

Source is approval-time attendance synchronization history, not present-day schedule or mutable calendar contents. Only recorded after-state Congé payé cells are counted. Repos/Férié were skipped by synchronization. Date/period is deduplicated across saved workspaces and overlapping approvals. Cross-year/month requests are clipped correctly. Pending/rejected/withdrawn/cancelled requests do not count; pending cancellation keeps original approval counted. Presentation-only request archiving does not erase duration. Manually marked calendar CP without an approved application is not included.

Old approvals without ledger/synchronization events or malformed snapshots produce an explicit incomplete flag and known subtotal, rather than silently inventing a duration. Legacy no-data requests are not backfilled, and no signed documents or calendars are modified. No database migration, background worker or external mail.

Tests: test_leave_totals.py synthetic real approval/cancellation flows, test_person_names.py HTTP owner/manager/Admin access and denied outsider, and the owned MySQL temporary-table workflow in test_responsable_workflow_mysql.py.

Validation: 68 focused SQLite/HTTP tests and 26 owned MySQL temporary-table workflows passed; frontend build/lint and 78 tests passed. Local API restarted. After user login, the Profile annual total and monthly days rendered correctly; switching 2026 to 2025 updated the year and totals without errors. Existing local account has zero approved recorded days; non-zero half-day and cancellation cases are verified by synthetic HTTP/MySQL tests. No production deployment.
