# ILCI · Espace RH 模块索引

产品显示名称统一为 ILCI · Espace RH，React 共用 `frontend/src/brand.ts`；图标为 RH，网页标题与图标在 `frontend/index.html` 和 `frontend/public/favicon.svg`。技术服务名、数据库表、环境变量与本地会话键沿用原标识，避免影响部署和已有登录。随 2026-10-23 已签归档版本上线。

## 入口与环境

| 环境 | 入口 | 服务 |
| --- | --- | --- |
| React 前端 | `frontend/src/main.tsx` → `App.tsx` → `AppWorkspace.tsx` | 本地 5173 |
| 线上 API | `backend/rebuild_server.py`，复用 `backend/main.py` | `presence-react-api`，8015 |
| 独立本地 API | `backend/local_rebuild.py` | 8013，独立 MySQL 3307；另有 [Responsable 合成测试账号](responsable-qa-account.md)，仅本地 5174 / 8014 与独立 QA 数据库 |
| 保留 API | `backend/main.py` 的 flat 导入布局 | API 8001；仍保留会话和权限保护 |
| 旧界面备案 | `archive/streamlit/source.zip`、Git 标签 `archive/streamlit-20261023` | 独立解压查阅；线上旧服务另行维护 |

线上使用原数据库及私有文件；本地副本位于被忽略的 `.local-test/rebuild/`，不能作为生产数据发布。当前发布及回退见 [发布索引](rebuild-release-readiness.md)，运行配置见 [后端说明](../backend/README.md)。当前依赖见 [架构说明](architecture.md)，旧界面恢复见[备案说明](../archive/streamlit/README.md)。

## 前端定位

路径均相对于 `frontend/src/`。

| 功能 | 入口与直接依赖 |
| --- | --- |
| 导航、页面加载、主题 | `AppWorkspace.tsx`；`workspacePages.ts` 管理标题/悬停说明；`workspaceToolbar.ts`、`components/WorkspaceToolbar.tsx` 管理当前页工具栏；`ThemeModeProvider.tsx`、`theme.ts`、`index.css` 管理深浅主题和布局 |
| 草稿保护 | `draftGuard.ts`、`draftRules.ts`、`components/DraftGuardProvider.tsx`；`components/DraftStatus.tsx` 在右下角显示紧凑悬浮状态，考勤与保存按钮合并；不占页面布局，离开确认保持；已打开页面保留状态，隐藏页不挂载工具栏 |
| 考勤 | `attendance/AttendancePage.tsx`；`AttendanceToolbar.tsx` / `PeriodToolbar.tsx` 顶栏；`FicheEditor.tsx` 编辑；`FicheExportMenu.tsx` 导出菜单；`AttendanceContractDialog.tsx` 复用 `admin/ContractEditor.tsx` 的表单补录缺失合同 |
| 日历 | `attendance/AttendanceCalendar.tsx`、`useCalendarSelection.ts`、`calendarStatus.ts`；半天拖选、批量应用及状态颜色 |
| 签署与批量结果 | `attendance/SignaturePreview.tsx`、`components/BatchExportResult.tsx`；下载和签署独立，异常提交不自动重发 |
| 已签文件 | `attendance/SignedFichesPanel.tsx`、`signedFichesApi.ts`；考勤顶部 Signatures 页签与个人 Profile 月度列表；每页 20 项、共用 DocumentPreview、管理员月度 ZIP；正式版仅显示正式文件，本地保留生产/SPECIMEN 切换 |
| Documents | `documents/DocumentsPage.tsx` 每页 20 条、跨页勾选；`DocumentsToolbar.tsx`；`DocumentUpload.tsx` 上传弹窗；`DocumentRow.tsx` 紧凑横条及操作；`documentSearch.ts` 搜索；`DocumentPreview.tsx` / `PdfPreview.tsx` 私有预览 |
| 员工名册 | `personnel/PersonnelPage.tsx`、`PersonnelToolbar.tsx`；`PersonnelEditor.tsx` / `PersonnelDetail.tsx`、`EmploymentEditor.tsx`、`PersonnelHistory.tsx`；`personnelValidation.ts` 共用字段标签及校验；`EmploymentFields.tsx` 集中合同与三类特殊身份表单；`RegistrationContract.tsx` 显式合同预填；`PersonnelInformation.tsx` 共用名册/个人资料展示及 EmploymentSummary；`SensitiveNumber.tsx` 受控短时查看；`PersonnelExport.tsx` Excel 导出确认 |
| Administration | `admin/AdminPage.tsx`；`UserEditor.tsx`、`OrganizationFields.tsx`、`ContractEditor.tsx`、`GroupEditor.tsx`；`CreateAccount.tsx` 建号弹窗；`AccountAccess.tsx` 登录状态；`AdvancedMaintenance.tsx` 资源维护 |
| 人员选择和姓名 | `personDisplayName.ts`；`admin/AccountOption.tsx` / `accountOptionLayout.ts` 分列显示；`components/DepartmentBadge.tsx` / `departmentColor.ts` 部门标签和颜色 |
| Mon équipe / 月度 Profile | `team/TeamPage.tsx`、`OrganizationCards.tsx`、`PersonProfile.tsx`；顶部年份/月下拉选择，`ProfileMonth.tsx` 仅展示所选月份，上半部为四项汇总、下半部为文档归档明细、请假日期/状态及签署记录；年度 API 数据供月份切换复用。已签文件复用 `DocumentPreview` 按需预览/下载，仍由后端鉴权。弹窗固定高度、内部纵向滚动，页面说明移入提示；此界面已于 2026-10-28 发布，见 [发布记录](releases/profile-month-ui-20261028.md) |
| 个人设置 / 名册关联 | `profile/ProfilePage.tsx` 的账号/员工资料页签；`AccountPersonnelSection.tsx` 复用关联卡与编辑窗口；`accountPersonnelApi.ts` 请求 |
| 仪表板 / 请假 | `dashboard/DashboardPage.tsx`、`PersonnelDistribution.tsx`、`InterviewEditor.tsx`；`LeaveWorkspace.tsx` 和 `components/Leave*` |
| 共用请求、日期、登录 | `api.ts`、`download.ts`、`components/FrenchDateField.tsx` / `components/CalendarMonthHeader.tsx`、`authSession.ts`、`components/PasswordRecovery.tsx`、`capabilities.ts` |

页面说明放在顶部标题 Tooltip 中。常规提示与操作使用主题蓝色，错误/删除保留语义红色。页签切换保留表单输入；个人设置只在账号页签显示保存操作。

侧边栏将员工名册与 Administration 集中在带边框的 `Espace Admin` 分组，仅 Admin/SuperAdmin 可见；进入这两个页面时顶部同步显示区域标识。考勤、团队等共享页面仍位于日常导航，接口权限保持由后端执行。

`personnel/PersonnelTable.tsx` 按 Excel 模板 B:R 的字段顺序提供 Informations de base、Emploi et dates、Situations particulières 三个网格页签；基础页在桌面同屏展示姓名、国籍、性别、地址、社保号、出生日期、职位与资格，按屏宽分配列宽；单元格尽量单行，长内容省略并在悬停时显示全文，性别缩写 H/F（不改变存储值）。分组表头固定、每人一行；任职与特殊身份显示最新任职段，历史通过完整档案查看。窄屏优先展示关键列，不增加横向滚动；社保号继续复用受控显示组件，切换页签即卸载并清除明文。Excel 导出仍使用完整筛选结果与全部任职段。

名册 Informations 页直接展示最新任职段的外籍、临时工、青年/学徒情况，并可打开 `EmploymentEditor.tsx` 的 situationsOnly 模式编辑；无任职段时引导填写首段，有多段时可进入完整期间页。人员类型、职位、资格仍由基本资料编辑器维护，期间身份保存复用原版本检查和审计接口。

`PersonnelDetail.tsx` 使用居中悬浮 Dialog，桌面最大宽度 lg、高度最多 820px，四周保留背景；内容区内部滚动，标题和底部操作固定，切换页签不随内容收缩。小屏同样保留窗口边距。

## 后端定位

路径均相对于 `backend/`。路由模块通过注册函数接收认证、数据库和授权依赖，不反向导入 `main.py`。

| 模块 | 职责 / 接口 |
| --- | --- |
| `main.py` | 认证、组织、账户、提醒、面谈、考勤 JSON 保存；`/login`、`/restore-session`、`/list-users`、`/groups`、`/dashboard`、`/get-config/{user_id}`、`/save-config/{user_id}` |
| `role_policy.py` | 角色继承、实时身份/会话检查、最后 SuperAdmin 保护、角色审计和用户名保留 |
| `login_guard.py` | 密码连续失败 5 次冷却 10 分钟；未满 5 次距上次失败满 10 分钟重置。users 行锁与 UTC 持久计数/时间，用户名/邮箱共用状态，HTTP 429 + Retry-After；成功清零。已上线 |
| `account_lifecycle.py` / `account_purge.py` | 停用/恢复、删除登录、状态历史；永久删除前检查业务依赖，禁止级联清理他人资料 |
| `account_names.py` / `person_names.py` | 建号必填姓/名、users 仅增列迁移；统一显示优先关联名册的结构化姓/名，再用账号姓/名、旧名册全名、邮箱前缀。不猜测旧姓名，不读 NIR 等隐私字段 |
| `account_onboarding.py` | 建号邮箱校验、后端安全随机 8 位大写字母＋数字、初始密码邮件（TLS）；仅在建号提交后尝试一次，返回发送结果、不保存或返回原密码。`account_lifecycle.py` 增列 `initial_password_pending`，`/leave/session` 返回建议改密标志；前端 `CreateAccount.tsx` 选择随机/手动，`AppWorkspace.tsx` 提醒修改但不限制访问。说明见 [建号通知](account-onboarding.md)，2026-10-27 已上线 |
| `personnel_register.py` / `personnel_rules.py` / `personnel_store.py` | 员工名册路由、校验、存储及审计；`/personnel-register*`、`/account-personnel/{username}`、`/choices`、`/link` |
| `personnel_sensitive.py` | 独立 keyring、AES-256-GCM、员工 ID 绑定和短时查看保护 |
| `personnel_export.py` | 按 `template/Exemple-de-modele-de-registre-unique-du-personnel.xlsx` 生成 Excel；复用名册筛选、全部分页、每段任职一行；`POST /personnel-register/export.xlsx` 仅有效管理员并要求敏感导出确认，包含 NIR 明文且先提交逐人审计，不在服务器保存文件 |
| `attendance_api.py` / `attendance_revision.py` | `/attendance/workspace`、`/save`、`/open/{username}`、`/calendar/{id}`、`/export`；事务内校验完整配置版本；本地负责人隔离历史表与合同补录见 [修复说明](responsable-attendance-fix.md) |
| `attendance_permissions.py` | 保存字段与历史 ID 核对；`attendance_scope` 隔离不可操作历史表，`preserve_hidden_sheets` 在完整版本锁内原样合并；Admin 归档，SuperAdmin 明确确认后才可删除 |
| `contract_history.py` / `personnel_contract_sync.py` | 合同期间、重叠校验、归档/恢复；`/users/{username}/contracts*`；`POST /attendance/contracts/{username}` 允许 Admin 或直接负责人补录缺少的当月合同并写入 `contract_setup_events`（2026-10-27 已上线）；关联名册按稳定 ID 同步入职与合同预计结束日期，实际离职独立；被考勤或名册引用的合同不可删除 |
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

2026-10-28 已发布修正：`main.update_group` 先锁定并确认部门存在，再执行更新；MySQL 同值更新的 `rowcount=0` 不再误报 404，缺失 ID 仍返回 404。`require_responsable_for_managed_groups` 统一要求建号／组织编辑的管理部门必须配合显式 Responsable 角色；`admin/OrganizationFields.tsx` 在没有该角色时禁用管理部门并提示，移除角色时清空管理部门草稿。此修正已随前后端同步发布，不自动更改线上账号权限。

| 需求 | 模块 |
| --- | --- |
| 考勤必填校验、命名、文件生成 | `fiche_exports.py`、`fiche_serialization.py` |
| 默认安排、半天状态、旧日历数据兼容 | `planning.py`、`calendar_view.py`；后者仅保留纯规则，旧渲染器与组件在备案源码包中 |
| 合同快照、月份有效区间和例外 | `contract_periods.py` |
| Excel / 实习生 Word | `ExcelGen.py` + `Fiche_Exemple.xlsx`；`DocxGen.py` + `template_stagiaire.docx` |
| ZIP 名称、容量与完整性 | `document_downloads.py` |
| 组织架构投影 | `organization_model.py` |

旧 UI 与其专属请求/展示模块已分离备案；考勤、合同、文件生成仍使用根目录共用逻辑。不能因界面迁移就删除 `vacances`、`absences`、`arret`、`calendar_overrides` 等兼容数据。当前运行依赖为根目录 `requirements.txt`，开发依赖为 `requirements-dev.txt`，旧版依赖保留在备案包内。

## 数据与权限约束

2026-10-28 账号设置与旧表签署修复：`GET /account` 只返回已认证本人公开账号资料，`profile/ProfilePage.tsx` 不再依赖 `/list-users` 的团队范围；纯 Responsable 无名册资料也能访问设置。`signature_flow.resolve_manager(employee_username=...)` 按旧表稳定关联账号鉴权并读取当前邮箱，拒绝改用其他邮箱身份；`attendance_signatures.py` 仅在本次预览/发送副本中补齐员工和负责人联系信息，保留其他字段及历史存储。预检先显示授权确认的收件人，再显示表单校验错误；未通过校验不提供可发送指纹。发送重新读取并核对指纹，邮箱、负责人或内容变化必须重新确认。不改合同日期，不自动创建名册；预览不触发实际发送，发送仍须显式确认。

- `users` 保存账户/角色/会话版本及独立的 `first_name`、`last_name`；新建账号前后端都必填姓/名，界面通过邮箱登录，未指定旧式用户名时后端生成稳定的随机账号标识。姓名格式为“名 + 空格 + 姓”，不按姓名建立身份关联。`personnel_register*` 独立保存员工资料、任职期间与审计。一对一关联使用唯一 `linked_username`，不从账号自动生成正式员工档案。
- `Presence.form_content` 保存绑定月份、合同快照和日历。历史表不因当前合同变化丢失；新建须有有效合同，已归档记录须先恢复。
- `personnel_contracts` 是考勤合同历史，员工名册通过稳定的 `account_contract_id` 同步入职与预计结束日期；`actual_exit_date` 独立维护，状态与 Excel Sortie 列使用实际离职。旧 `exit_date` 原值保留、不自动推断实际离职。
- 登记规则更新：新员工的姓、名、国籍、性别、出生日期、地址、职位、资格、人员类型、初始入职日期必填。`registration_schema_version=2` 的新档案后续编辑也不能清空必填项；旧档案维持版本 1，可逐步补全、不拆分旧姓名。三类特殊身份按任职期间保存，勾选后后端强制校验对应字段；仅增列迁移，已随账号姓名版本发布。
- `transport_receipts` 保存元数据；文件留在私有目录，下载须鉴权。ZIP 最多 100 份、100 MiB，失败不返回部分归档。
- `leave_requests` / `leave_events` 保存申请和操作历史；不自动联动考勤、余额或邮件。
- SuperAdmin 继承 Admin；Admin 可查看、新建、修改、归档/恢复，永久删除仅 SuperAdmin，且须满足归档/停用、精确确认和依赖检查。权限最终由后端执行。
- 负责人在授权团队内查看/填考勤。普通员工和经理可上传自己的文件；Admin 可代传。非 Admin 经理的 Documents 归档仍受只读规则限制。
- 全员名册仅 Admin/SuperAdmin；普通账户只能读取本人关联资料。直接负责人查看下属月度 Profile 不获得名册地址、生日或 NIR。
- 停用和删除登录保留业务历史；状态/角色变化撤销旧会话，用户名保留防止旧身份继承。已发出的外部签署链接须在 ClawShow 单独处理。
- NIR 生产已显式启用，仅 Admin/SuperAdmin 可录入及受审计查看；本地继续使用合成号码。密钥独立备份，普通 DTO 不含原号/密文，查看禁止缓存。详见 [名册说明](personnel-register.md) 和 [启用记录](releases/nir-enable-20261022.md)。
- 签署回执只代表发送被受理。新增完成回调和最终 PDF 归档仅处理配置回调后的新请求；`all_signed` 完成后后台回收，原始已签 PDF 不重写，按发送时考勤月份归档。两张 signed_fiche 表和永久目录须同时备份；旧签署不自动补发或补收。2026-10-23 已上线，真实 Sandbox 双签回存已验证；配置及安全边界见 backend/README.md。
- 2026-10-28 已发布：`signature_flow.resolve_manager(include_identity=True)` 提供当次实际负责人账号，`esign.py` 在发送前写入 `signed_fiche_documents.manager_username`。`signed_fiche_store.signed_fiche_scope` / `can_view_signed_fiche` 统一列表、Profile 与文件访问：Admin/SuperAdmin 或员工本人可查完整历史；Responsable 必须同时是当前直接负责人和该次签署负责人。发送操作者不等于签署负责人。旧表只新增可空列，不根据当前组织或 sender 猜测历史归属；未知归属旧记录仅对负责人隐藏，PDF、回调凭据、员工/Admin 历史均保留。账号改名同步身份列，不改原 PDF。
- 2026-10-28 已发布：`admin/organizationValidation.ts` 在组织保存确认前提示缺少直接负责人或管理部门；`UserEditor.tsx` 把接口保存错误显示于确认弹窗内部，失败保留窗口，成功关闭，避免错误被遮罩挡住。

## 维护与记录

先读对应入口及直接依赖；共用逻辑修改时检查全部调用方。保留历史快照、月份绑定、草稿和鉴权，不清空数据。

- 部署前对比本地/服务器差异，备份源码、数据库、私有文件和签署 SQLite；同时保护仍可访问的旧 API。
- 2026-10-23 API 与归档验证：[发布记录](releases/signed-archive-20261023.md)；最新前端：[正式版隐藏测试入口](releases/production-signatures-ui-20261023.md)。
- 运行配置、启动和安全回退：[后端说明](../backend/README.md)。
- 旧发布时间线：[历史发布记录](releases/20261014-20261021.md)。
- 本地快照及启动：[本地重构说明](local-rebuild.md)；`scripts/start_local_rebuild.ps1`。
- 本地签署入口：`backend/pdf_test_gateway.py` 仅开放限时 PDF 和有大小限制的 `/esign/callback`，回调固定转发至本地 8013、由私有 API 核对随机凭据；不开放其他接口。`scripts/start_pdf_test_tunnel.ps1` 配置两条公网地址并重启本地 API，`stop_pdf_test_tunnel.ps1` 清除并关闭；`restart_local_rebuild.ps1` 只重启本工作区 8013。共用 `local_process.ps1` 核对监听进程与 Windows venv 启动器身份，拒绝停止其他工作区进程。`local_signing_policy.py` 限制指定真实收件人和时效，不部署本地网关/私有配置。

2026-10-23 旧版分离与名册整理：[整理记录](source-cleanup.md)。备案基线为 `de055e8`，本地名册开发进度在拆分前单独提交 `ff1fc16`；未改变生产部署或业务数据。

## 本轮回归补充（2026-10-23）

`personnel_rules.entry_precedes_birth` 统一检查入职与出生日期；名册编辑和账户合同同步共用。首次登记通过 `sync_account_contracts(..., initial_period_id=...)` 明确传入初始任职，仅唯一对应时采用账户合同日期，避免留下重复任职；常规同步继续按稳定 ID 执行。登记异常中的 NIR 在模型级校验前包装为 SecretStr；日期自定义标题显式固定法语。测试范围与限制见 [精简后回归](cleanup-regression-20261023.md)。

2026-10-23 账号姓名：创建表单移除手工账号标识，改为必填 Prénoms/Nom；账号姓名直接保存在 users，无需自动创建正式员工档案。关联名册的结构化姓名优先，旧记录不自动拆分。模块兼容 flat/package 导入及旧数据库缺列读取；发布时需同步仍可访问的保留 API。存储核对与范围见 [账号姓名说明](account-names.md)。

当前生产前端：`profile-month-ui-20261028T115926Z` / `3ce60e6eac9f`，Profile 使用月份下拉与单月详情；API 保持 `contact-access-20261028T095545Z` / `f0609e84168b`，保留本人设置读取及旧工时表当前签署邮箱修复。静态目录与 API current 指针分别维护，见 [本次界面发布记录](releases/profile-month-ui-20261028.md) 和 [API 发布记录](releases/account-contact-20261028.md)。

前端随后独立更新为 `admin-workspace-ui-20261023T151905Z` / `63ef6c3935f2`，Admin 导航集中分组并显示区域标识；API 保持上述版本。静态目录、核对与独立界面回退见 [Admin 分组发布记录](releases/admin-workspace-ui-20261023.md)。

最新前端为 `ux-polish-ui-20261023T174314Z` / `4ce4a8710f49`，发布紧凑草稿提示、日历选择同行提示、签署窗口精简、模板式名册页签、详情悬浮窗口和简化登录文案；API 保持账号姓名版本。核对范围、静态目录和逐层回退见 [界面整理发布记录](releases/ux-polish-ui-20261023.md)。

## Responsable 流程验证（2026-10-26）

`test_responsable_workflow_mysql.py` 使用本机 3307、连接级 MySQL TEMPORARY 表、合成人员与真实 JWT/数据库角色校验，串联建号、名册关联、合同、负责人创建工时表、半天修改、保存与重新打开；SMTP 替换为测试 mock，不修改永久业务表。需显式设置 `PRESENCE_MYSQL_TEMP_TEST=1`。与负责人权限、历史月份、版本保护、名册关联及登记回归合计 63 项，在本地源码与最近发布 API 提交 `b6ca83d5e046` 对应源码均通过；前端纯逻辑测试 46 项通过。未登录线上 Robin 账号核验实际数据，也未进行浏览器端完整验收或真实签署/邮件发送。

已复现的业务边界：登记名册不自动建号；关联账号后，仅填入职日期且未设 CDI/合同预计结束日期，不一定生成考勤合同，团队可见仍可能因当月无有效合同返回 409。完整关联任职日期可同步创建合同，也可由 Admin 在 Administration 维护。考勤工作区目前仍按操作者账号保存，Admin 和负责人不会自动共享同一份草稿；若保存的工作区包含已调出管理范围的人员，读取会明确返回 409 并保留原数据，不会静默删除历史。这些限制已在新增测试中固定记录，不能将本轮通过等同于线上个案已修复。

2026-10-27 线上只读个案排查：使用当前部署源码的账户读取/权限函数与考勤路由，在独立进程中连接只读事务，禁用保存回调，未导入主应用或执行启动迁移。确认管理权限正常时，旧工作区缺少 `account_username` 的记录仍会在 `require_employee` 中判为不可访问，继而由 workspace 全量一致性检查返回 409、阻止整个编辑页加载；目标账号没有 `personnel_contracts` 记录时，新建另返回缺少当月合同的 409。线上日志存在对应 workspace 409；没有写入生产资料、重启服务或发布。新增 MySQL 临时表测试固定“名单可见＋历史未关联表＋目标无合同”组合，与权限/历史回归共 18 项通过，属于故障复现而非修复完成。
