# ILCI · Espace RH 模块索引

产品显示名称统一为 ILCI · Espace RH，React 共用 `frontend/src/brand.ts`；图标为 RH，网页标题与图标在 `frontend/index.html` 和 `frontend/public/favicon.svg`。技术服务名、数据库表、环境变量与本地会话键沿用原标识，避免影响部署和已有登录。随 2026-10-23 已签归档版本上线。

## 入口与环境

| 环境 | 入口 | 服务 |
| --- | --- | --- |
| React 前端 | `frontend/src/main.tsx` → `App.tsx` → `AppWorkspace.tsx` | 本地 5173 |
| 线上 API | `backend/rebuild_server.py`，复用 `backend/main.py` | `presence-react-api`，8015 |
| 独立本地 API | `backend/local_rebuild.py` | 8013，独立 MySQL 3307 |
| 保留版 | `online_test.py` / `pages/`；flat `main:app` | Streamlit 8501 / API 8001 |

线上使用原数据库及私有文件；本地副本位于被忽略的 `.local-test/rebuild/`，不能作为生产数据发布。当前发布及回退见 [发布索引](rebuild-release-readiness.md)，运行配置见 [后端说明](../backend/README.md)。旧版调用关系见 [架构说明](architecture.md)。

## 前端定位

路径均相对于 `frontend/src/`。

| 功能 | 入口与直接依赖 |
| --- | --- |
| 导航、页面加载、主题 | `AppWorkspace.tsx`；`workspacePages.ts` 管理标题/悬停说明；`workspaceToolbar.ts`、`components/WorkspaceToolbar.tsx` 管理当前页工具栏；`ThemeModeProvider.tsx`、`theme.ts`、`index.css` 管理深浅主题和布局 |
| 草稿保护 | `draftGuard.ts`、`draftRules.ts`、`components/DraftGuardProvider.tsx`；已打开页面保留状态，隐藏页不挂载工具栏 |
| 考勤 | `attendance/AttendancePage.tsx`；`AttendanceToolbar.tsx` / `PeriodToolbar.tsx` 顶栏；`FicheEditor.tsx` 编辑；`FicheExportMenu.tsx` 导出菜单 |
| 日历 | `attendance/AttendanceCalendar.tsx`、`useCalendarSelection.ts`、`calendarStatus.ts`；半天拖选、批量应用及状态颜色 |
| 签署与批量结果 | `attendance/SignaturePreview.tsx`、`components/BatchExportResult.tsx`；下载和签署独立，异常提交不自动重发 |
| 已签文件 | `attendance/SignedFichesPanel.tsx`、`signedFichesApi.ts`；考勤顶部 Signatures 页签与个人 Profile 月度列表；每页 20 项、共用 DocumentPreview、管理员月度 ZIP；正式版仅显示正式文件，本地保留生产/SPECIMEN 切换 |
| Documents | `documents/DocumentsPage.tsx` 每页 20 条、跨页勾选；`DocumentsToolbar.tsx`；`DocumentUpload.tsx` 上传弹窗；`DocumentRow.tsx` 紧凑横条及操作；`documentSearch.ts` 搜索；`DocumentPreview.tsx` / `PdfPreview.tsx` 私有预览 |
| 员工名册 | `personnel/PersonnelPage.tsx`、`PersonnelToolbar.tsx`；`PersonnelEditor.tsx` / `PersonnelDetail.tsx`、`EmploymentEditor.tsx`、`PersonnelHistory.tsx`；`personnelValidation.ts` 注册必填及期间校验；`EmploymentFields.tsx` / `ContractFields.tsx` / `SpecialEmploymentFields.tsx` 任职与三类特殊身份；`RegistrationContract.tsx` 显式合同预填，`EmploymentSummary.tsx` 只读情况；`SensitiveNumber.tsx` 受控短时查看；`PersonnelExport.tsx` Excel 导出确认 |
| Administration | `admin/AdminPage.tsx`；`UserEditor.tsx`、`OrganizationFields.tsx`、`ContractEditor.tsx`、`GroupEditor.tsx`；`CreateAccount.tsx` 建号弹窗；`AccountAccess.tsx` 登录状态；`AdvancedMaintenance.tsx` 资源维护 |
| 人员选择和姓名 | `personDisplayName.ts`；`admin/AccountOption.tsx` / `accountOptionLayout.ts` 分列显示；`components/DepartmentBadge.tsx` / `departmentColor.ts` 部门标签和颜色 |
| Mon équipe / 月度 Profile | `team/TeamPage.tsx`、`OrganizationCards.tsx`、`PersonProfile.tsx`；名单/架构页签，月度考勤、文件、请假和签署发送回执 |
| 个人设置 / 名册关联 | `profile/ProfilePage.tsx` 的账号/员工资料页签；`AccountPersonnelSection.tsx` 复用关联卡与编辑窗口；`accountPersonnelApi.ts` 请求 |
| 仪表板 / 请假 | `dashboard/DashboardPage.tsx`、`PersonnelDistribution.tsx`、`InterviewEditor.tsx`；`LeaveWorkspace.tsx` 和 `components/Leave*` |
| 共用请求、日期、登录 | `api.ts`、`download.ts`、`components/FrenchDateField.tsx` / `components/CalendarMonthHeader.tsx`、`authSession.ts`、`components/PasswordRecovery.tsx`、`capabilities.ts` |

页面说明放在顶部标题 Tooltip 中。常规提示与操作使用主题蓝色，错误/删除保留语义红色。页签切换保留表单输入；个人设置只在账号页签显示保存操作。

名册 Informations 页直接展示最新任职段的外籍、临时工、青年/学徒情况，并可打开 `EmploymentEditor.tsx` 的 situationsOnly 模式编辑；无任职段时引导填写首段，有多段时可进入完整期间页。人员类型、职位、资格仍由基本资料编辑器维护，期间身份保存复用原版本检查和审计接口。

## 后端定位

路径均相对于 `backend/`。路由模块通过注册函数接收认证、数据库和授权依赖，不反向导入 `main.py`。

| 模块 | 职责 / 接口 |
| --- | --- |
| `main.py` | 认证、组织、账户、提醒、面谈、考勤 JSON 保存；`/login`、`/restore-session`、`/list-users`、`/groups`、`/dashboard`、`/get-config/{user_id}`、`/save-config/{user_id}` |
| `role_policy.py` | 角色继承、实时身份/会话检查、最后 SuperAdmin 保护、角色审计和用户名保留 |
| `login_guard.py` | 密码连续失败 5 次冷却 10 分钟；未满 5 次距上次失败满 10 分钟重置。users 行锁与 UTC 持久计数/时间，用户名/邮箱共用状态，HTTP 429 + Retry-After；成功清零。已上线 |
| `account_lifecycle.py` / `account_purge.py` | 停用/恢复、删除登录、状态历史；永久删除前检查业务依赖，禁止级联清理他人资料 |
| `person_names.py` | 明确关联名册姓名，缺失时使用有效邮箱前缀；不按同名/邮箱猜测身份，不读 NIR 等隐私字段 |
| `personnel_register.py` / `personnel_rules.py` / `personnel_store.py` | 员工名册路由、校验、存储及审计；`/personnel-register*`、`/account-personnel/{username}`、`/choices`、`/link` |
| `personnel_sensitive.py` | 独立 keyring、AES-256-GCM、员工 ID 绑定和短时查看保护 |
| `personnel_export.py` | 按 `template/Exemple-de-modele-de-registre-unique-du-personnel.xlsx` 生成 Excel；复用名册筛选、全部分页、每段任职一行；`POST /personnel-register/export.xlsx` 仅有效管理员并要求敏感导出确认，包含 NIR 明文且先提交逐人审计，不在服务器保存文件 |
| `attendance_api.py` / `attendance_revision.py` | `/attendance/workspace`、`/save`、`/open/{username}`、`/calendar/{id}`、`/export`；事务内校验配置版本 |
| `attendance_permissions.py` | 保存字段与历史 ID 核对；Admin 归档，SuperAdmin 明确确认后才可删除 |
| `contract_history.py` / `personnel_contract_sync.py` | 合同期间、重叠校验、归档/恢复；`/users/{username}/contracts*`；关联名册按稳定 ID 同步入职与合同预计结束日期，实际离职独立（本地新版）；被考勤或名册引用的合同不可删除 |
| `documents.py` / `document_numbers.py` | `/transport-receipts*` 上传、搜索数据、下载、ZIP、改月、归档和受控删除；按巴黎上传日分配 `YYYYMMDD-0001`，删除不回收编号 |
| `esign.py` / `signature_flow.py` / `signature_placement.py` | PDF 转换、当前直接负责人解析、双签顺序和签名定位；`/fiches/pdf`、`/send-fiche`、`/esign/documents/{token}` |
| `attendance_signatures.py` / `signature_attempts.py` / `server_signing_readiness.py` | React 预检/发送、幂等回执、配置就绪检查；`/attendance/signature-preview`、`/signature-send` |
| `signed_fiches.py` / `signed_fiche_store.py` | 新请求随机凭据回调 `/esign/callback`、持久下载队列、最终 PDF 私有归档、逐文件鉴权/完整性和月度 ZIP；`/signed-fiches*`。main 共用注册，flat/package 兼容；已上线 |
| `person_profiles.py` | `/person-profiles/{username}?year=YYYY`：本人、当前直接负责人或 Admin 的月度汇总 |
| `team_directory.py` | `/organization-directory` 和 `/my-team`：公共架构与本人部门目录 |
| `leave_requests.py` / `leave_rules.py` / `leave_store.py` | `/leave/session`、`/leave-requests*`；半天、审批前撤回、当前直接负责人审批、Admin 代审及历史 |
| `resource_maintenance.py` / `destructive_audit.py` | `/maintenance/resources/{kind}`：部门、请假、面谈的归档/删除；独立最小删除审计 |
| `ui_capabilities.py` | `/ui-capabilities`：公开布尔能力与环境信息，不替代业务鉴权 |

旧 API 使用 flat 导入，共享模块须兼容；仅供 React 入口使用的模块可保持包导入。改名时由 `main.py` 直接调用 `personnel_store.rename_personnel_account_link` 和 `leave_store.rename_leave_owner`，与账户更新共用事务。

## 根目录共用逻辑

| 需求 | 模块 |
| --- | --- |
| 考勤必填校验、命名、文件生成 | `fiche_exports.py`、`fiche_serialization.py` |
| 默认安排、半天状态、旧日历兼容 | `planning.py`、`calendar_view.py`；旧交互组件 `calendar_selector_component/` |
| 合同快照、月份有效区间和例外 | `contract_periods.py` |
| Excel / 实习生 Word | `ExcelGen.py` + `Fiche_Exemple.xlsx`；`DocxGen.py` + `template_stagiaire.docx` |
| ZIP 名称、容量与完整性 | `document_downloads.py` |
| 组织架构投影 | `organization_model.py` |

旧 UI 和相关共用模块仍用于回退。不能因 React 不显示旧输入区就删除 `vacances`、`absences`、`arret`、`calendar_overrides` 等兼容数据。

## 数据与权限约束

- `users` 保存账户/角色/会话版本；`personnel_register*` 独立保存员工资料、任职期间与审计。一对一关联使用唯一 `linked_username`，不从账号自动生成正式员工档案。
- `Presence.form_content` 保存绑定月份、合同快照和日历。历史表不因当前合同变化丢失；新建须有有效合同，已归档记录须先恢复。
- `personnel_contracts` 是考勤合同历史，员工名册通过稳定的 `account_contract_id` 同步入职与预计结束日期；本地新版的 `actual_exit_date` 独立维护，状态与 Excel Sortie 列使用实际离职。旧 `exit_date` 原值保留、不自动推断实际离职。
- 本地登记规则更新：新员工的姓、名、国籍、性别、出生日期、地址、职位、资格、人员类型、初始入职日期必填。`registration_schema_version=2` 的新档案后续编辑也不能清空必填项；旧档案维持版本 1，可逐步补全、不拆分旧姓名。三类特殊身份按任职期间保存，勾选后后端强制校验对应字段；仅增列迁移，尚未部署。
- `transport_receipts` 保存元数据；文件留在私有目录，下载须鉴权。ZIP 最多 100 份、100 MiB，失败不返回部分归档。
- `leave_requests` / `leave_events` 保存申请和操作历史；不自动联动考勤、余额或邮件。
- SuperAdmin 继承 Admin；Admin 可查看、新建、修改、归档/恢复，永久删除仅 SuperAdmin，且须满足归档/停用、精确确认和依赖检查。权限最终由后端执行。
- 负责人在授权团队内查看/填考勤。普通员工和经理可上传自己的文件；Admin 可代传。非 Admin 经理的 Documents 归档仍受只读规则限制。
- 全员名册仅 Admin/SuperAdmin；普通账户只能读取本人关联资料。直接负责人查看下属月度 Profile 不获得名册地址、生日或 NIR。
- 停用和删除登录保留业务历史；状态/角色变化撤销旧会话，用户名保留防止旧身份继承。已发出的外部签署链接须在 ClawShow 单独处理。
- NIR 生产已显式启用，仅 Admin/SuperAdmin 可录入及受审计查看；本地继续使用合成号码。密钥独立备份，普通 DTO 不含原号/密文，查看禁止缓存。详见 [名册说明](personnel-register.md) 和 [启用记录](releases/nir-enable-20261022.md)。
- 签署回执只代表发送被受理。新增完成回调和最终 PDF 归档仅处理配置回调后的新请求；`all_signed` 完成后后台回收，原始已签 PDF 不重写，按发送时考勤月份归档。两张 signed_fiche 表和永久目录须同时备份；旧签署不自动补发或补收。2026-10-23 已上线，真实 Sandbox 双签回存已验证；配置及安全边界见 backend/README.md。

## 维护与记录

先读对应入口及直接依赖；共用逻辑修改时检查全部调用方。保留历史快照、月份绑定、草稿和鉴权，不清空数据。

- 部署前对比本地/服务器差异，备份源码、数据库、私有文件和签署 SQLite；同时保护仍可访问的旧 API。
- 2026-10-23 API 与归档验证：[发布记录](releases/signed-archive-20261023.md)；最新前端：[正式版隐藏测试入口](releases/production-signatures-ui-20261023.md)。
- 运行配置、启动和安全回退：[后端说明](../backend/README.md)。
- 旧发布时间线：[历史发布记录](releases/20261014-20261021.md)。
- 本地快照及启动：[本地重构说明](local-rebuild.md)；`scripts/start_local_rebuild.ps1`。
- 本地签署入口：`backend/pdf_test_gateway.py` 仅开放限时 PDF 和有大小限制的 `/esign/callback`，回调固定转发至本地 8013、由私有 API 核对随机凭据；不开放其他接口。`scripts/start_pdf_test_tunnel.ps1` 配置两条公网地址并重启本地 API，`stop_pdf_test_tunnel.ps1` 清除并关闭；`restart_local_rebuild.ps1` 只重启本工作区 8013。共用 `local_process.ps1` 核对监听进程与 Windows venv 启动器身份，拒绝停止其他工作区进程。`local_signing_policy.py` 限制指定真实收件人和时效，不部署本地网关/私有配置。

本次清理仅移除无调用的旧级联清理函数、失效导入与转导入；历史实现可从 Git `7ed872e` 查询。
