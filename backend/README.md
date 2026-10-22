# API Presence

## Current release: 2026-10-22

`personnel-profile-20261022T125536Z` deploys the personnel register, account/register
associations, monthly profiles, account lifecycle controls, shared SuperAdmin/Admin
archive policy, name projections and compact React workspace. Both API entries
(8015 and retained 8001) enforce the updated authentication and resource guards.
The retained API uses flat modules; shared modules must support flat imports as
well as `backend.*` imports. New legacy files must be readable by `streamlitapp`,
not merely root; verify imports with the actual service identity before restart.

The application symlink points to
`/var/www/presence-react/releases/personnel-profile-20261022T125536Z`.
Source/configuration, database, private documents and the signing-attempt SQLite
backup are under root-only
`/var/backups/presence-react/personnel-profile-20261022T125536Z`.
Production NIR remains disabled; no local snapshot, local encryption key or local
SuperAdmin promotion is imported. Existing production roles/passwords are retained.

Safe UI rollback:
`sudo /var/backups/presence-react/personnel-profile-20261022T125536Z/rollback-ui.sh`.
It restores the previous static interface and keeps the updated API guards and
all business data. Do not restore the old API blindly after accounts have been
disabled or login-deleted: it would bypass their new access-state checks. Database
restoration is a separate, deliberate recovery operation, never the default.
See [release verification](../doc/releases/personnel-profile-20261022.md).

React remake released 2026-10-15 as `react-20261015T101334Z`. Public UI: `https://app.presence.217.65.145.103.nip.io/`. Service `presence-react-api` runs the full-repository entry `backend.rebuild_server:app` on loopback 8015. Legacy services remain active for rollback. See `../doc/rebuild-release-readiness.md` for backup paths and the guarded rollback command.

This directory contains the version-controlled FastAPI service used by the Presence interface.

## Deployment

## Paid leave module (React service)

The React interface in `../frontend` registers `/leave/session` and `/leave-requests`
through `leave_requests.py`. Deploy `leave_rules.py`, `leave_store.py`, and
`leave_requests.py` with the new `main.py`; startup creates the two leave tables.
The entry point also updates request ownership on account rename and detaches it
on account deletion, preserving immutable operation history. Back up both source
and MariaDB before restarting, and verify authorization and transactions first.

The retired fictitious-account preview was removed on 2026-10-16. Use
`local_rebuild.py` with the isolated authorized data snapshot for local work.
The production React entry remains `rebuild_server.py`, reusing `main.py`.

## Production environment

The production service reads its credentials from `/etc/presence-app/presence-app.env`, via the systemd drop-in for `presence-app.service`. Do not commit that file or any password, SMTP credential, database credential, or signing secret.

## Signature électronique

`esign.py` converts attendance sheets to PDF and requests an employee signature
followed by the responsible manager through ClawShow eSign. `signature_flow.py`
resolves and validates the two people. Both signatures and their placement were
verified in a completed sandbox employee PDF; the user subsequently confirmed the
live employee-to-manager signing flow. The complete modular version was deployed
on 2026-10-14. The intern end-to-end signing flow remains to be verified.
The production environment must provide `CLAWSHOW_ESIGN_API_KEY`;
it must only exist in `/etc/presence-app/presence-app.env`. LibreOffice is required
on the API server for the Excel/Word-to-PDF conversion. Generated PDFs are kept in
`PRESENCE_ESIGN_DOCUMENTS_DIR` (default: `/var/lib/presence-app/esign-documents`)
behind an unguessable URL for 24 hours, so ClawShow can retrieve the document.
`PRESENCE_ESIGN_PUBLIC_URL` must be the public FastAPI origin (not the Streamlit
interface URL), for example `https://presence.example.org`.

## Documents

Transport receipts are stored privately in `PRESENCE_TRANSPORT_RECEIPTS_DIR`
(default: `/var/lib/presence-app/transport-receipts`). The API accepts one PDF,
JPG or PNG up to 10 MB at a time. The database stores only file metadata,
assignment, document type, bound month/year and archive state; downloads remain protected by the authenticated API.
At startup, the `transport_receipts` table is created if it does not exist.
`document_numbers.py` adds a nullable unique `document_number` and a
`document_daily_sequences` counter table. It fills missing references in upload
order (`created_at`, then `id`), using the timestamp converted to Europe/Paris.
New uploads reserve `YYYYMMDD-0001` references in the same transaction as the
metadata insert, using a per-day InnoDB row lock. Archive/month edits retain the
reference; deletion never decrements the counter. At least four sequence digits
are displayed, extending naturally past 9999. Single downloads and ZIP entries
include the reference; private storage filenames and original names stay intact.
Deploy `document_numbers.py` with `documents.py` for both active API entries.
Rollback may retain the additive column, unique index and counters, avoiding any
loss of issued references; do not restore a database dump over later uploads.

Each authenticated employee or manager may upload only for their own username.
For non-admins, `/transport-receipts/assignees` returns only that username and the
upload route rejects a different target with HTTP 403 before writing files or
database rows. Admin retains upload-on-behalf permission. Managers retain team
read access, without additional archive, period-correction or deletion rights.
Existing document ownership is unchanged.

`documents.py` owns the existing `/transport-receipts` routes. The name stays for
URL and client compatibility. Frontend ZIP exports use these individual authenticated
downloads, enforce a 100 MB limit, and do not return incomplete archives.

Before replacing the deployed API, take both a source backup and a MariaDB dump. At startup, the API performs only additive schema migrations: it adds optional organization columns to `users` and creates the group tables if they do not already exist. It does not delete existing user accounts or `Presence` records.

The expected command is:

```text
uvicorn main:app --host 127.0.0.1 --port 8001
```

The deployed service uses `/var/www/presence-app/main.py`. For the flat server layout,
deploy `main.py`, `documents.py`, `esign.py`, `contract_history.py`,
`attendance_permissions.py`, `signature_flow.py` and `signature_placement.py` together in that directory.
The team-directory update also requires `team_directory.py` with the updated
`main.py`. It adds authenticated read-only organization and own-department directory
routes; existing personnel, contract and Documents authorization remains unchanged.
Local package imports (`backend.main`) and flat imports (`main`) are both supported.
Route registration injects the original authentication and database dependencies;
submodules must not import `main.py` back.

Signature placement requires `pypdf` from `requirements.txt`. The current ClawShow
signing-page renderer maps field coordinates through a 595 by 842 reference page.
`signature_placement.py` finds the labels in the final PDF and sends normalized
explicit coordinates, without mixing anchors and explicit page/x/y fields. This
avoids the observed landscape-anchor shift. Recheck this mapping if the provider
changes its field coordinate convention. Rotated/cropped or ambiguous templates
are rejected rather than signed at a guessed location.

Frontend deployment must include the changed pages and their common helpers:
`api_client.py`, `fiche_exports.py`, `fiche_serialization.py`, `planning.py`,
`calendar_view.py`, `document_ui.py` and `document_downloads.py`, alongside existing
export/template and period modules. Compare with production before overwriting files;
the server may contain changes absent locally. Replacing only a page or `main.py`
can leave its new imports missing.

The attendance batch-selection update also requires `fiche_batch_ui.py` together
with `pages/Fiches.py`. Batch download and signing are separate actions using the
same explicit selection. Signing retains the employee's registered direct manager.
Completion webhook processing and automatic signed-PDF archival are not implemented.

The module map is in [../doc/modules.md](../doc/modules.md). Current refactoring and
verification status is recorded in [../doc/refactoring.md](../doc/refactoring.md).

## Runtime configuration

Configure values in the protected server environment; the table lists names only.
Do not copy that environment into the repository or frontend secrets.

| Variables | Purpose |
| --- | --- |
| `PRESENCE_API_TOKEN` | Required session signing secret |
| `PRESENCE_DB_HOST`, `PRESENCE_DB_USER`, `PRESENCE_DB_PASSWORD`, `PRESENCE_DB_DATABASE` | Required MariaDB connection |
| `PRESENCE_SMTP_SERVER`, `PRESENCE_SMTP_PORT`, `PRESENCE_SMTP_USER`, `PRESENCE_SMTP_PASSWORD` | Required mail configuration |
| `PRESENCE_APP_URL` | Required frontend origin for account emails |
| `PRESENCE_TRANSPORT_RECEIPTS_DIR` | Private long-term Documents storage; default given above |
| `PRESENCE_ESIGN_DOCUMENTS_DIR` | Temporary signing PDFs; default given above |
| `PRESENCE_ESIGN_PUBLIC_URL` | Public HTTPS API origin from which ClawShow fetches PDFs |
| `CLAWSHOW_ESIGN_API_KEY`, `CLAWSHOW_ESIGN_NAMESPACE` | Provider key; namespace defaults to `ilci` |
| `PRESENCE_ESIGN_ORGANIZATION_NAME`, `PRESENCE_ESIGN_ORGANIZATION_FULL_NAME` | Mail institution identity; defaults to ILCI and its full name |
| `PRESENCE_ESIGN_REPRESENTATIVE_NAME`, `PRESENCE_ESIGN_REPRESENTATIVE_TITLE` | Mail closing name and position; defaults to resolved manager and `Responsable` |

The frontend reads `URL_PRESENCE` from its own `.streamlit/secrets.toml`. A local
UI can connect to the production API, so actions there affect real server data.
To run only the local UI, use the project virtual environment:

```powershell
.\.venv\Scripts\python.exe -m streamlit run online_test.py --server.address 127.0.0.1 --server.port 8502 --server.headless true
```

## Release and rollback

2026-10-21 email login patch: `/login` continues accepting the `username` JSON
field, now resolving either username or bound email (case-insensitive, surrounding
identifier whitespace ignored) to exactly one canonical username. Ambiguous
cross-account matches fail with the same 401 as invalid credentials. Password
whitespace is significant. Session and remember tokens use the canonical username.
Both live APIs received only this login resolution patch against their reviewed
source digests; no schema, passwords or business data were changed. The UI was
built from deployed commit 4d6e0ad with only the login field updated. Local account
lifecycle/profile work remains separate and is not deployed by this patch.
Release: `email-login-20261021T151000Z`. Root-only source/database backup and guarded
combined UI/API rollback: `/var/backups/presence-react/email-login-20261021T151000Z-backend/rollback.sh`.
Rollback restores code and UI only, never the database dump. Email ownership
verification, OTP and login throttling are not added by this patch.

The 2026-10-14 modular release backup is
`/var/backups/presence-interface/online-release-20261014T091212Z`. It contains a
file manifest, the prior deployed source files and a MariaDB dump. Runtime keys
were preserved in place and are not part of that source release payload.

For each future release, compare source differences, back up every replaced file
and the database, validate flat backend imports, and deploy all required modules.
Keep existing file ownership and modes. Restart both services, verify their active
state, public API routes and Streamlit health, and inspect startup errors.

For a source rollback, stop the services, restore the manifest's existing files
from that release backup, remove only files marked newly added by that release,
and restart and verify the services. Leave uploaded Documents and server secrets
intact. Additive schema changes do not require automatic database restoration;
restoring a database dump can discard later writes and needs a separate decision.

Health checks establish startup and reachability. They do not establish successful
login, uploads, export contents or email delivery. Known behavior needing correction,
including destructive account deletion and private session caches, is tracked in
[the maintenance plan](../doc/code-audit.md).


## React server entry (released 2026-10-15)

`backend.rebuild_server:app` is the server entry for the full React migration.
Run from the complete repository root; retain root-level generators and templates.
It reuses `main.app`, authentication, current permissions and checked saves, then
registers the React attendance and signature adapters. Do not deploy the isolated
`local_rebuild.py` entry or its runtime/snapshot files.

The entry uses the existing server environment and private document paths. Set
`PRESENCE_SIGNATURE_ATTEMPTS_DB` to a writable persistent SQLite file (default
`/var/lib/presence-app/signature-attempts.sqlite3`) and back it up with the service.
Readiness validates the selected signing key prefix and public HTTPS PDF origin;
it does not prove external PDF accessibility or email delivery. No provider call
is made by this readiness check. Signing receipt storage is not final PDF archival.

`ui_capabilities.py` registers public `GET /ui-capabilities`, returning only
boolean feature availability and environment type. Local startup explicitly
reports account creation and mail as disabled and blocks the corresponding routes.
The capability endpoint does not replace authorization on business endpoints.

Uninstalled systemd/Nginx examples are in `../deploy/`. They use a separate API
port and preserve the old service for rollback. Review the actual server paths,
service user, proxy prefixes, SSL host, permissions and environment before use.
The new entry still inherits the original startup migrations and reminder worker;
starting it against production is a release action, requiring backup and approval.
See `../doc/rebuild-release-readiness.md` for acceptance and remaining boundaries.

The new service disables its reminder worker using `PRESENCE_CONTRACT_REMINDER_WORKER=0` while the retained legacy API owns the worker. If retiring the legacy API later, enable this worker on the new service. Manual reminders remain authorized normally. Source rollback switches the UI and stops the new service; it does not restore the database dump or remove uploaded files.


## 人员姓名显示（2026-10-22，本地，未部署）

### 账号 Profile 的名册关联（本地，未部署）

Admin/SuperAdmin 在账号详情或 Profile 的「Données du registre du personnel」中选择「Associer une fiche / Modifier l’association」，可手动关联已有登记、更换关联或选择无关联。复用唯一的 `personnel_register.linked_username`，不自动创建员工，不猜测身份。普通账号可在个人设置和本人月度 Profile 只读查看自己的资料；直接负责人仅查看下属月度概览不会获得地址、生日等名册信息。社保号/密文及其存在标记不进入这些 Profile 响应。

- `GET /account-personnel/{username}`：重查有效会话，允许本人或 Admin；仅选取普通资料及任职期间，响应 no-store。
- `GET /account-personnel/{username}/choices`：仅 Admin，搜索未归档且未被其他账号占用的记录，返回最小选择信息。
- `PUT /account-personnel/{username}/link`：仅 Admin；请求包含 `employee_id`、`expected_employee_id`、`expected_version`、`target_version`（无关联时相应值为 null）。当前关联和两个员工版本须与读取时相同；保留唯一约束并检查目标是否被其他账号占用。users 排序锁后再按员工 ID 排序锁，解除原关联、建立新关联与操作审计同事务提交，失败整体回滚。

关联不会覆盖名册内容、合同或考勤快照，解除也不会删除资料。修改关联后刷新姓名/账号/团队等缓存；普通账号不能通过直接调用接口修改关联或查询全员名册。原账号邮箱/密码设置不受影响。本轮仅静态检查和构建，未运行功能回归或操作真实关联，未上线。

账号选择菜单现分为姓名/邮箱、职位/人员身份、管理权限三列，Admin / Super Admin 使用简写。只有 Admin/SuperAdmin 的 `/list-users` 额外返回可选 `job_title`，按名册 `linked_username` 精确关联；其他角色、公共人员目录、会话 DTO 不增加职位或敏感字段。`main.admin_job_titles` 与纯姓名投影分开，旧库缺名册表时空职位兼容。

考勤打开已补历史快照边界：即使当前合同已改日期或归档，明确关联同一账号、同一月份且未归档的已有考勤仍可打开；不重绑/改写其合同快照。不存在历史表时仍要求有效合同；归档考勤不会被自动恢复。前端显示该人员已有月份及切换入口，账号选择会激活该人员当月已有页签。名册日期与考勤合同独立，实际日期仍由管理员确认更正，本轮未代改。

此轮完整后端 358 项中 352 通过、6 项 opt-in 跳过，新增 `test_attendance_open_history.py` 9 项与 `test_admin_job_projection.py` 8 项。前端 43 项通过，完整 lint、类型检查、生产构建通过。重启本地 API 后以只读 GET 验证新增职位投影与历史考勤仍在；未执行真实合同编辑、考勤生成、邮件、签署或部署。

人员显示取自 `personnel_register.full_name` 与 `linked_username` 的明确关联；没有关联姓名时使用有效邮箱的前缀，缺邮箱时使用中性提示。需要补姓名时在 `Registre du personnel` 更正姓名并关联对应账号，不会按同名或邮箱自动合并身份。`person_names.py` 提供批量公开姓名投影，只有姓名/关联列与必要邮箱进入该逻辑；不读取 NIR 等其他名册隐私字段。后端旧安装没有名册表时仍可用邮箱前缀显示。

此变更不更名账号或修改权限，不回写既有考勤、合同或审计快照。界面和签署人的姓名来自当前关联资料；历史考勤中的姓名输入仍代表该文件的导出内容。内部用户名继续用于关联/鉴权，因此界面显示名不是新的身份凭证。账号删除确认支持当前邮箱并兼容旧用户名；不会自动填写确认。当前只在本地更新，未部署。

2026-10-22 功能回归与修复：

- Python 全量 discover：341 项中 335 通过、6 项显式 opt-in 跳过；另设置进程环境 `PRESENCE_MYSQL_TEMP_TEST=1` 执行 `test_account_lifecycle_mysql`、`test_personnel_sensitive_mysql`，6 项全部通过。MySQL 测试固定本地 3307，并且只写连接级 TEMPORARY 表；不修改真实账号或员工资料。
- 新增 `test_person_names.py` 27 项和 `test_person_name_workflows.py` 18 项，覆盖姓名/邮箱回退、明确关联、同名不同账号访问隔离、名册/Profile/请假权限、社保号及凭据不外泄、历史数据不改写；账号删除仍要求精确邮箱或技术 ID，姓名不能充当确认。
- 修复后端/组织模型对异常邮箱的回退规则，与前端保持单个 `@` 且无内部空白；下拉框增加邮箱辅助识别同名人员。
- 请假历史在账号更名后，当前账号无法解析时保留已有姓名快照的显示，拒绝原始账号 ID、空值和异常邮箱；存储的 actor ID/事件行不改变。其他没有可靠姓名快照的历史操作仍使用 `Ancien compte`，不按姓名猜测关联。
- 前端 35 项规则测试、TypeScript、全量 lint、生产构建通过。9 个真实本地只读 API 及 6 个认证拒绝检查通过；未执行实际上传、邮件、签署或部署。浏览器视觉/点击交互未由自动化覆盖。

可复现命令：`.venv/Scripts/python.exe -W ignore::DeprecationWarning -m unittest discover -p 'test*.py'`；前端在 `frontend/` 运行 `npm test`、`npm run lint -- src test`、`npm run build`。

## 本地角色与归档升级（2026-10-22，未部署）

SuperAdmin 继承 Admin 的业务能力，并独享受控永久删除；普通 Admin 继续查看、新建、修改、归档和恢复。普通 Admin 不能给自己或他人分配 SuperAdmin，也不能修改 SuperAdmin 的权限或登录状态。后端每次敏感写入在 users 排序锁之后重查角色、账号状态及会话版本；自身及最后有效 SuperAdmin 有保护。角色变更会使旧会话和记住登录凭据失效。

本次仅在本地副本为现有 `ChristopheH` 赋权，使用 `scripts/promote_local_superadmin.py` 显式核对固定 loopback 数据库、账号及邮箱。不会自动创建超级账号或在服务器启动时赋权。该测试身份暂时可见，正式隐藏以后另行开发。角色备份仅包含原角色元数据，位于被忽略的 `.local-test/rebuild/role-backups/`；如回退角色，仍须递增 `session_version`，不可回写旧版本使旧令牌重新有效。

永久删除要求先归档/停用并精确确认。账号仍有人员、考勤、合同、文件、请假或访谈关联时拒绝物理删除，应分别处理关联资料；不会级联删除其替别人上传的文件。保留独立最小操作审计和被删除/更名账号的用户名占位，避免旧令牌在重用用户名后复活。合同被历史考勤引用时拒删。文档数据库删除已提交但物理文件清理失败时返回 `cleanup_pending`，需要后续清理，不得当作文件完全擦除。

`resource_maintenance.py` 在两个 React API 入口注册并增量迁移归档列及独立审计表；`/maintenance/resources/{kind}` 管理部门、请假与访谈。普通 Admin 可归档/恢复，SuperAdmin 可永久删除；待审批请假须先处理，部门须先解除成员关联，自动生成的访谈须在账号停用后才可物理删除。名册、Documents、合同和考勤沿用各自接口，详细边界见 [模块索引](../doc/modules.md)。生产尚未部署本次迁移。

## 员工名册 NIR 保护（2026-10-22，仅本地，未部署）

仅在本地使用合成号码测试，未部署、未向真实员工档案写入 NIR。`PRESENCE_PERSONNEL_NIR_ENABLED` 默认关闭，只有精确为 `1` 且 `PRESENCE_PERSONNEL_NIR_KEYS_FILE` 指向的独立 AES-256-GCM keyring 有效时开启；不派生会话密钥、不在服务器入口自动生成生产 key。本地入口首次 provision 被忽略的 `.local-test/rebuild/keys/personnel-nir.json`，已有文件不覆盖。POSIX 首次创建为 0600；Windows fresh 环境须另配 NTFS ACL（本机本次已配置），mode 参数不能实施 ACL。密钥须独立备份，轮换保留仍由密文或备份引用的旧 key。接口、短时查看/审计、加密范围及专项/全量验证结果见 [员工名册说明](../doc/personnel-register.md)。

## 账号停用改版（2026-10-21，仅本地，未部署）

- `account_lifecycle.py` 的增量迁移增加 `users.is_active`（默认 1）、`users.session_version`（默认 0）与 `account_status_events`，不删除业务数据。现有零日期默认值会触发部分 MySQL 严格模式的 DDL 失败；迁移仅在当前连接的新增 DDL 期间去掉 NO_ZERO_DATE/NO_ZERO_IN_DATE，并在 finally 恢复原 session sql_mode，不修改全局配置或旧日期值。
- 前端改用激活接口；旧 `/delete-user/{username}` 必须返回 409。Admin 停用自己被拒绝。状态变化与审计记录在同一事务提交，按用户名排序锁定账户，与角色更改使用相同锁顺序；锁后再次读取操作者的活跃状态、Admin 权限及会话版本。
- 登录、持久会话恢复、每次鉴权均检查状态；会话携带版本，旧会话按 0 兼容。停用及重新启用均递增版本、清空记住登录和重置令牌，重新启用不会恢复旧登录。密码保留。停用不会撤回已发给 ClawShow 的签署链接，外部待签请求需要在 ClawShow 另行处理。
- 上线前按本文原有流程备份数据库和所有拟替换源码；分别比对新 React API 与仍运行的旧 API。**所有仍可访问的认证入口、旧永久删除接口都必须同步保护**，再开放前端停用功能，否则旧 API 可能绕过停用限制。不能整体覆盖有未同步差异的服务器 main.py。
- 回退时保留新增表和列及已发生的状态记录，不恢复数据库旧备份。不得直接运行不识别 is_active/session_version 的旧后端：它会重新放行停用账号和旧会话。必要的 UI 回退可以保留新认证/停用保护并禁用旧删除入口；后端回退须保留这些补丁。
- 本地已启动增量迁移并只读查看 Admin 确认窗口；未实际修改任何账号状态，未运行功能测试/真实邮件/签署，未部署服务器。


### 后续：停用后删除登录（2026-10-21，本地）

用户确认删除登录而保留历史资料。新增 `users.login_deleted_at`（默认 NULL），状态日志增加 `action`（默认 activation，兼容旧记录）。新 DELETE `/users/{username}/login` 必须携带 `{ "confirmation": "完整用户名" }`，仅活跃 Admin 可操作；先锁定并重新验权，要求目标已停用、不是本人。删除使用身份墓碑：只更新登录标记、随机替换旧密码哈希、递增会话版本并清除登录/重置令牌，不 DELETE 任何业务行或 unlink 文件；用户名/邮箱保持占用，避免后续新账号取得原档案。重复删除幂等，不能通过 activation 重新启用。保留事件操作者、时间和 login_deleted 动作；历史资料仍可经授权查询。

部署需要沿用上节全 API 一致性与备份要求；回退补丁还须识别 login_deleted_at，禁止恢复已删除身份的登录或 activation。不能把回退到不识别墓碑字段的旧代码当作安全回退。当前仅构建/lint/Python 静态检查与本地启动迁移；未实际删除任何账号，未运行删除功能回归或发布。


### 账号流程测试入口（2026-10-21）

- `.venv/Scripts/python.exe -m unittest test_account_lifecycle -v`：24 项隔离测试，不连 MySQL。
- `.venv/Scripts/python.exe -m unittest discover -v`：默认跳过显式 opt-in 的 MySQL 测试。
- 在本地 MySQL 3307 正常运行、私有 `.local-test/rebuild/runtime.json` 存在时，设置进程环境 `PRESENCE_MYSQL_TEMP_TEST=1` 后运行 discover；这次实际执行 179 项全部通过。连接固定 127.0.0.1:3307，数据库名必须 presence_local_rebuild，使用连接级 TEMPORARY 表覆盖测试连接内的 users/events，真实账户不变，不创建数据库或持久测试员工。连接关闭即清理；测试代理拒绝业务表写入。
- 覆盖先停用再删除、精确用户名确认、普通员工/经理越权、自我停用、会话版本/旧令牌兼容、记住登录撤销、删除不可重新启用、重复请求、历史及他人文件保留、审计失败时事务回滚。MySQL 测试使用真实 SQL/哈希/事务；并发部分目前是 SQLite 事务模拟（25 次重复通过），未做线上或 MySQL 多连接压力测试。
- 首轮失败来自旧测试模拟和 SQL 参数断言，已随新 Profile/编号/签署代码更新。当前改动仍未发布，现有本地员工和线上账号未实际停用或删除。
