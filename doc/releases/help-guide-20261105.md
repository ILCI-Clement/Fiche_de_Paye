# French help guide — 2026-11-05

Published frontend: `rh-help-guide-20261105T101144Z`. Nginx serves this release's frontend/dist; sole API/current remains `rh-mailnotify-20261104T162306Z` and was not restarted.

Sidebar-bottom question-mark button opens a French step-by-step guide on the current module. Topics appear in a vertical left tab list, with independently scrollable detail content on the right. Switching topics resets content scroll, preserves the underlying page and drafts, and does not trigger business actions. Shared role policy filters manager/admin topics. Desktop/sidebar and mobile drawer use one shared guide dialog. The account-administration hover description now reflects unified personnel/account creation.

Verification: 77 frontend tests, lint and production build passed. Nine public assets and all 248 staged resource hashes matched; 457 prior release files and the API PID were unchanged. Three private-route guards and 20 authenticated account/profile read checks passed. No business writes, emails or signing requests. Backend/mail workers remain unchanged.

Protected backup: `/var/backups/presence-react/rh-help-guide-20261105T101144Z`, with guarded `rollback-ui.sh` to restore the previous Nginx static root only. Backend/data unchanged. Keep the difference between the frontend release and API/current path when preparing the next deployment.
