# ILCI · RH 模块索引

产品显示名称统一为 ILCI · RH，React 共用 `frontend/src/brand.ts`；图标为 RH，网页标题与图标在 `frontend/index.html` 和 `frontend/public/favicon.svg`。技术服务名、数据库表、环境变量与本地会话键沿用原标识，避免影响部署和已有登录。随 2026-10-23 已签归档版本上线。

## 入口与环境

| 环境 | 入口 | 服务 |
| --- | --- | --- |
| React 前端 | `frontend/src/main.tsx` → `App.tsx` → `AppWorkspace.tsx` | 本地 5173 |
| 线上 API | `backend/rebuild_server.py`，复用 `backend/main.py` | `presence-react-api`，8015 |
| 独立本地 API | `backend/local_rebuild.py` | 8013，独立 MySQL 3307；另有 [Responsable 合成测试账号](responsable-qa-account.md)，仅本地 5174 / 8014 与独立 QA 数据库 |
| Retired services | `presence-app` / `presence-interface` | Masked since 2026-10-30; ports 8001 / 8501 closed; protected archive only |
| 旧界面备案 | `archive/streamlit/source.zip`、Git 标签 `archive/streamlit-20261023` | 独立解压查阅；线上旧服务另行维护 |

线上使用原数据库及私有文件；本地副本位于被忽略的 `.local-test/rebuild/`，不能作为生产数据发布。当前发布及回退见 [发布索引](rebuild-release-readiness.md)，运行配置见 [后端说明](../backend/README.md)。当前依赖见 [架构说明](architecture.md)，旧界面恢复见[备案说明](../archive/streamlit/README.md)。

## 前端定位

路径均相对于 `frontend/src/`。

| 功能 | 入口与直接依赖 |
| --- | --- |
| 导航、页面加载、主题 | `AppWorkspace.tsx`；`workspacePages.ts` 管理标题/悬停说明；`workspaceToolbar.ts`、`components/WorkspaceToolbar.tsx` 管理当前页工具栏；`ThemeModeProvider.tsx`、`theme.ts`、`index.css` 管理深浅主题和布局 |
| 草稿保护 | `draftGuard.ts`、`draftRules.ts`、`components/DraftGuardProvider.tsx`、`components/DraftStatus.tsx` 管理当前页草稿提示与离开确认；`AppWorkspace.tsx` 仅挂载当前模块，确认放弃后卸载旧表单并清除草稿登记，已保存数据仍从 API/查询缓存读取，后台忙碌操作继续阻止切页 |
| 考勤 | `attendance/AttendancePage.tsx`；`AttendanceToolbar.tsx` / `PeriodToolbar.tsx` 顶栏；`FicheEditor.tsx` 编辑；`FicheExportMenu.tsx` 导出菜单；`AttendanceContractDialog.tsx` 复用 `admin/ContractEditor.tsx` 的表单补录缺失合同 |
| 日历 | `attendance/AttendanceCalendar.tsx`、`useCalendarSelection.ts`、`calendarStatus.ts`；半天拖选、批量应用及状态颜色 |
| 签署与批量结果 | `attendance/SignaturePreview.tsx`、`components/BatchExportResult.tsx`；下载和签署独立，异常提交不自动重发 |
| 已签文件 | `attendance/SignedFichesPanel.tsx`、`signedFichesApi.ts`；考勤顶部 Signatures 页签与个人 Profile 月度列表；每页 20 项、共用 DocumentPreview、管理员月度 ZIP；正式版仅显示正式文件，本地保留生产/SPECIMEN 切换 |
| Documents | `documents/DocumentsPage.tsx` 每页 20 条、跨页勾选；`DocumentsToolbar.tsx`；`DocumentUpload.tsx` 上传弹窗；`DocumentRow.tsx` 紧凑横条及操作；`documentSearch.ts` 搜索；`DocumentPreview.tsx` / `PdfPreview.tsx` 私有预览。`CategoryManager.tsx` 管理类型，`DocumentTypeFields.tsx` 共享选择，`DocumentClassification.tsx` 管理员更正已有分类，`DocumentTypeBadge.tsx` 统一圆角描边和彩色填充，`documentCategory.ts` 兼容历史显示；自定义类型2026-10-28 已上线，见 [说明](document-categories.md) |
| 员工名册 | `personnel/PersonnelPage.tsx`、`PersonnelToolbar.tsx`；`PersonnelEditor.tsx` / `PersonnelDetail.tsx`、`EmploymentEditor.tsx`、`PersonnelHistory.tsx`；`personnelValidation.ts` 共用字段标签及校验；`EmploymentFields.tsx` 集中合同与三类特殊身份表单；`RegistrationContract.tsx` 显式合同预填；`PersonnelInformation.tsx` 共用名册/个人资料展示及 EmploymentSummary；`SensitiveNumber.tsx` 受控短时查看；`PersonnelExport.tsx` Excel 导出确认 |
| Administration | `admin/AdminPage.tsx`；`UserEditor.tsx`、`OrganizationFields.tsx`、`ContractEditor.tsx`、`GroupEditor.tsx`；建号统一由 `personnel/PersonnelEditor.tsx` 进入；`AccountAccess.tsx` 登录状态；`AdvancedMaintenance.tsx` 资源维护 |
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

2026-10-28 已上线：`document_categories.py` 管理 `/document-categories*` 的已登录读取、实时 Admin 管理；稳定代码复用原发票字段，兼容标准类型及旧记录；`documents.py` 在上传时校验类型，提供分类更正、变更审计与并发保护。启动幂等新增分类及审计两张表，见 [Documents 分类说明](document-categories.md)。

2026-10-28 已上线：`dashboard/WorkCenterPanel.tsx` 汇总权限范围内的请假、签署、PDF 回收与资料待办；`useWorkCenter.ts` / `workCenterApi.ts` 复用查询，`ReadinessStatus.tsx` / `AccountReadiness.tsx` 在首页、团队和账号页显示就绪状态。`workspaceNavigation.ts` 经草稿保护定位目标记录；管理账号默认首页改为 Dashboard。后端 `work_center.py` 注册只读 `GET /dashboard/work-center`，复用实时角色、考勤可管理范围、直属请假审批与签署历史隔离；名册完整度仅 Admin 可见，不作为考勤前置条件。`signed_fiche_store.py` 使用精确二进制 ID 关联以兼容混合排序规则。范围、测试与发布注意见 [待办中心说明](work-center.md)。

路径均相对于 `backend/`。路由模块通过注册函数接收认证、数据库和授权依赖，不反向导入 `main.py`。

| 模块 | 职责 / 接口 |
| --- | --- |
| `main.py` | 认证、组织、账户、提醒、面谈、考勤 JSON 保存；`/login`、`/restore-session`、`/list-users`、`/groups`、`/dashboard`、`/get-config/{user_id}`、`/save-config/{user_id}` |
| `role_policy.py` | 角色继承、实时身份/会话检查、最后 SuperAdmin 保护、角色审计和用户名保留 |
| `login_guard.py` | 密码连续失败 5 次冷却 10 分钟；未满 5 次距上次失败满 10 分钟重置。users 行锁与 UTC 持久计数/时间，用户名/邮箱共用状态，HTTP 429 + Retry-After；成功清零。已上线 |
| `account_lifecycle.py` / `account_purge.py` | 停用/恢复、删除登录、状态历史；永久删除前检查业务依赖，禁止级联清理他人资料 |
| `account_names.py` / `person_names.py` | 建号必填姓/名、users 仅增列迁移；统一显示优先关联名册的结构化姓/名，再用账号姓/名、旧名册全名、邮箱前缀。不猜测旧姓名，不读 NIR 等隐私字段 |
| `account_onboarding.py` | 建号邮箱校验、后端安全随机 8 位大写字母＋数字、初始密码邮件（TLS）；仅在建号提交后尝试一次，返回发送结果、不保存或返回原密码。`account_lifecycle.py` 增列 `initial_password_pending`，`/leave/session` 返回建议改密标志；当前前端由员工注册/外部建号统一生成临时密码，`AppWorkspace.tsx` 提醒修改但不限制访问。原独立弹窗见 `archive/react/`；后端兼容手动密码接口。说明见 [建号通知](account-onboarding.md)，2026-10-27 已上线 |
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

2026-10-29：Administration 的 `UserEditor.tsx` 复用当前月份 work-center 缺项，分别在名册、组织和合同折叠标题前显示 `components/CompletionDot.tsx` 黄点；悬停或聚焦显示原因。组织草稿同时复用 `organizationValidation`；未返回就绪结果时不推断名册或合同缺失。`AccountPersonnelSection` 接受可选的 `completionIssues`，不增加接口或权限。

UI motion: components/DialogMotion.tsx supplies the theme default Dialog enter/exit zoom and respects reduced motion. index.css adds button press/rebound and page fade-in. Page unmount and draft cleanup remain immediate; no old forms are retained for exit animations.

注册窗口外部人员分支：`PersonnelEditor.tsx` 的外部选项仅展示姓名、邮箱和 `RegistrationAccountFields`，调用共享 `adminApi.create` / 原 `/create-user`，不调用名册创建接口；不写合同或社保号。Administration 独立创建按钮暂时移除，已有账号管理保持。外部账号权限与其他账号同样由后端执行；账号专用流程的零名册/零合同测试位于 `test_registration_account.py`。

部门管理：Administration 工具栏的“Gérer les départements”按需打开 `GroupEditor.tsx` 弹窗，不再作为账号状态页面底部的常驻折叠块。弹窗关闭及部门切换只检查部门表单自身草稿，保存期间禁止关闭；接口与权限不变。

员工历史快照（本地）：`backend/personnel_snapshots.py` 管理三张独立 snapshot 表、Admin 下载/手动/重试接口、SuperAdmin 周期版本校验与 worker；`snapshot_schedule.py` 负责法国时间自然周期，`snapshot_crypto.py` 使用独立密钥整份加密 Excel。`personnel_export.generate_personnel_excel` 新增可选空表及快照时间元数据，不改变即时导出默认行为。`rebuild_server.py` / `local_rebuild.py` 注册服务；首页失败待办转入 `WorkDestination.snapshots`。前端 `personnel/PersonnelSnapshots.tsx` / `snapshotApi.ts` 提供名册顶部独立窗口；无删除接口，不重写成功快照。完整范围、密钥和恢复配置见 [历史快照](personnel-snapshots.md)。

账号状态访问：`GET /list-users` 默认仅返回正常登录账号；`include_inactive=true` 仅 Admin/SuperAdmin 可用，供 Administration 状态筛选读取历史。`admin/accountStatus.ts` 区分 Normal、Désactivé、Connexion supprimée；团队列表与人员指派不展示停用/删除登录账号。Administration 保留正常/停用/登录已删除/全部入口；PersonnelToolbar 的名册 Normal/Archivé/Tous 筛选直接显示在工具栏。账号停用不自动归档名册，归档详情与历史 API 不受此列表筛选影响。`AccountAccess.tsx` 明示“Supprimer l’accès”与“Supprimer définitivement…”；后端删除权限、先停用要求和历史依赖阻断规则不变。

2026-10-29：员工登记可选同时建号。`PersonnelEditor.tsx` / `RegistrationAccountFields.tsx` 复用组织权限控件；邮箱无账号时选填、建号时必填。`personnel_rules.EmployeeCreate.account` 接收组织配置，与已有账号关联互斥；`personnel_register.py` 在同一事务调用注入的 `main.insert_account`，保存名册、账号、组织审计与合同关联，提交后才调用 `welcome_created_account` 发送一次临时密码邮件。账号权限校验复用原建号逻辑。`personnel_store.py` 增量增加名册 `email` 列，不回填历史邮箱，旧客户端编辑不清空未提交的邮箱。邮箱不计入登记必填规则，也不自动改写已有登录账号邮箱。`backend/local_rebuild.py` 开放本地建号并强制 welcome mail 返回 False，线上配置不变；`ui_capabilities.py` 默认仍禁止普通快照环境建号，需本地入口显式开启。回归见 `test_registration_account.py`，事务集成只写连接级 MySQL TEMPORARY 表。
Signed archive annual view (local): AttendanceToolbar uses year-only PeriodSelector in the Signé tab; SignedFichesPanel groups the paginated annual list by month (newest first), with monthly ZIP actions. GET /signed-fiches accepts optional month; profile monthly queries and backend signature scope remain unchanged.
Published 2026-10-29: unified registration, account status tabs, encrypted register snapshots and annual/month-collapse signed archive view. See [release verification and rollback](releases/rh-unified-20261029.md). Production snapshot cycle is six months; first scheduled run 2027-01-01 03:00 Europe/Paris; no initial employee capture or simulation import.
Daily catalogs: backend/person_visibility.py annotates linked employee archival state without changing authentication. Default account list, team/organization, attendance choices, upload assignees and readiness hide inactive/deleted/archived people; Admin history retains access and an Archives tab. Register writes invalidate those frontend caches. See releases/person-visibility-20261029.md.
Team hierarchy: list-users includes a Responsable actor's direct reports even when those reports are themselves Responsable-only accounts; attendance/account edit authorization is unchanged. team/teamPeople.ts keeps employees and managers in the team table while honoring inactive/archived/deleted filtering. See releases/manager-hierarchy-20261029.md.
Manager attendance: role_policy.direct_business_report allows active non-Admin direct reports with Employe or Responsable function; main.user_can_manage and missing-contract setup share it. main.require_workspace_access keeps another manager's whole workspace private; personal attendance is created in the superior's own workspace. Supersedes the attendance restriction described in the earlier hierarchy release. See releases/manager-attendance-20261029.md.

Local account-role UI and visibility: admin/OrganizationFields.tsx separates business function from Admin/Super Admin checkboxes; admin/accountRoles.ts preserves the other category when editing. backend/person_visibility.account_list_visible enforces owner-only Super Admin catalog visibility, including list-users, team/organization, attendance, assignees, readiness and linked register rows/searches; existing historical records and encrypted snapshots are retained. See [scope and verification](account-role-visibility.md).

Account header access controls: admin/SystemAccessFields.tsx supplies shared Admin/Super Admin checkboxes with a confirmation dialog before changing the draft. UserEditor.tsx places them to the right of the account name and hides the duplicate controls in OrganizationFields; registration reuses the same confirmation component. Confirming the checkbox changes only the draft; existing organization save, role-change confirmation and backend authorization remain required.

Signed archive deletion (local): backend/signed_fiches.py exposes DELETE /signed-fiches/{archive_id} for freshly authorized Super Admin only, ready PDFs only, with exact-ID confirmation, independent destructive_audit signed_fiche event and guarded local file cleanup. SignedFichesPanel.tsx reuses PermanentDeleteDialog and refreshes lists/profile/work-center caches; external ClawShow originals and signature audit events are retained. See [scope and tests](signed-fiche-deletion.md).

Register display name (local): the menu, page heading, accessible labels and frontend references now use Utilisateurs (no Registre prefix); contextual prose uses liste des utilisateurs. The personnel route key, API paths, employee data model and Excel template remain unchanged.

Dashboard completion navigation (local): admin/accountSelection.ts resolves explicit username targets without falling back to the first account; AdminPage shows an inaccessible-target warning instead. backend/work_center.person_readiness labels missing manager contacts with the manager's name and supplies no hidden Super Admin target; dashboard/ReadinessStatus labels that action Ouvrir le responsable, while employee-specific corrections retain the employee's stable username. Regression coverage: test_work_center.py and frontend/test/accountSelection.test.mjs.

Register export formats (local): backend/personnel_pdf.py generates landscape A3 PDFs from the same template cells as Excel. personnel_register export.pdf allows Admin; export.xlsx now requires Super Admin. personnel_snapshots downloads default PDF, with Excel only for Super Admin; historical encrypted blobs stay unchanged. PersonnelExport and PersonnelSnapshots provide matching format choices, and PersonnelHistory displays PDF export audit events. Install the ReportLab dependency before deployment. See [permissions and validation](personnel-export-formats.md).

Published 2026-10-30 (Paris): the account-role/owner-only visibility controls, header confirmation UI, signed PDF deletion, Utilisateurs display name, dashboard target fix and PDF/Excel export permissions described above are now deployed. ReportLab 4.5.1 is installed. See [release verification, backups and recovery](releases/access-export-20261030.md).


## 请假取消（本地开发）

`backend/leave_requests.py` 新增 `POST /leave-requests/{id}/cancellation` 和 `/cancellation-decision`。未审批申请沿用本人撤回；已批准申请通过独立 `cancellation_status` 发起取消，批准前仍为有效 approved，请求被拒绝仍保留原批准状态，批准取消后 status=cancelled。当前直接负责人或其他 Admin 可审批取消，禁止本人自批；事务行锁与 leave_events 保留完整轨迹。`leave_store.py` 启动时仅追加 cancellation_status 列，不改写旧记录。

`components/LeaveDetail.tsx` 提供申请/审批取消及确认，`LeaveWorkspace.tsx` 待处理包含取消请求并增加已取消筛选；`work_center.py` 首页待办同步显示取消申请。现有请假流程不自动修改考勤、余额或发送邮件，本次取消遵循同样边界。测试：`test_leave_workflow.py`、`test_work_center.py`。

Leave confirmation UX (local): components/LeaveDetail.tsx uses a compact centered confirmation Dialog for decisions, withdrawals and cancellation requests/decisions. Confirmation actions stay outside the scrolling request history; Revenir, Escape or backdrop dismiss only the confirmation, and actions are disabled while submitting.

Leave summary navigation (local): LeaveWorkspace statistic cards are keyboard-accessible buttons selecting pending, approved or all within the current mine/team scope and resetting pagination to page 1; active filter uses aria-pressed and a highlighted border.


## Leave / attendance integration (local, 2026-10-30)

`backend/leave_attendance.py` owns approval preview, atomic calendar synchronization, per-half-day restoration ledger and signed-source guards. `attendance_sheet.py` is the shared month/contract sheet constructor used by manual creation and leave approval. `leave_store.py` adds `leave_attendance_cells` without rewriting historical records. Leave decisions and attendance writes commit together; missing contracts, protected sources and stale previews fail without partial approval. Explicit, fingerprint-bound confirmation is required before replacing sickness/absence/other non-work states. Overlapping approved leave transfers the original restoration baseline rather than removing another approved leave.

`main.save_config_checked` guards signed content and applies approved leave when a new workspace copy is created. `attendance_signatures.signed_source_lock` serializes re-reading, PDF generation and provider acceptance with calendar/leave writes; signed PDFs are never rewritten. `attendance_api.calendar_data` provides display-only pending-leave flags and read-only protection. `components/LeaveAttendanceCheck.tsx` shows the compact approval/cancellation synchronization preview, blockers and explicit conflict confirmation; `LeaveDetail.tsx` uses it inside the centered confirmation dialog. `LeaveWorkspace.tsx` invalidates attendance/profile caches after decisions. See [business rules and test scope](leave-attendance-sync.md). This supersedes the earlier local note that leave decisions have no attendance side effects; mail and leave-balance side effects remain absent.

Utilisateurs pagination (local): PersonnelPage.tsx renders row count, page-size selection and page navigation in a footer at the bottom of the registry panel, outside its scrolling table. PersonnelToolbar.tsx keeps only header actions and filters; pagination/filter state and server querying are unchanged.

Annual signed PDF download (local): signed_fiches.export.zip accepts an optional month; omitting it exports the selected year, all pages, for freshly authorized Admin/Super Admin only. SignedFichesPanel exposes Télécharger l’année in the annual view. Annual ZIP naming is ILCI_RH_YYYY.zip, containing ILCI_RH_YYYY/MM/original-signed-filenames; SPECIMEN exports remain separately prefixed. Monthly paths/naming, integrity checks, complete-archive requirement, 5,000-file/100-MiB bounds and private no-store responses are retained. Per-document audit uses yearly_export for annual downloads.

## File batch actions (2026-10-30, local)

`backend/file_batch.py` shares validation and postcommit cleanup. `documents.py` and `signed_fiches.py` provide `/batch` endpoints; `components/FileBatchDialog.tsx` confirms a frozen selection. Signé Normal / Archives uses additive `is_archived` metadata, without changing PDF contents. See [file batch actions](file-batch-actions.md).

## Shared header layout (2026-10-30, local)

`AppWorkspace.tsx` places the theme toggle and environment chip above the sidebar user row, including the mobile drawer; redundant sidebar prose is removed. `index.css` aligns page titles and portal toolbars in one desktop row with a common 72px minimum height. Crowded/narrow toolbars wrap without horizontal page overflow; the existing ResizeObserver still measures actual header height for Documents and personnel workspaces. Verified the main desktop pages at 1920px and settings at 1092px, crowded modules at 1280px, and the mobile drawer at 390px.

Published 2026-10-30 (Europe/Paris): `rh-file-batch-20261030T131557Z` now includes leave cancellation/attendance synchronization, annual signed ZIP, file batch actions, Utilisateurs pagination and the shared sidebar/header layout. See [release verification](releases/file-batch-layout-20261030.md). Production backup and guarded rollback are documented there; no local simulation data was imported.

Administration account list (local, 2026-10-30): `admin/AccountList.tsx` renders aligned clickable account rows using the existing AccountOption presentation. `AdminPage.tsx` keeps status/search filters while opening UserEditor in a bounded, scrollable Dialog; closing uses the existing draft/busy guard and returns to the same list. An explicit unavailable navigation target shows a warning and never falls back to another account. Account mutations and backend authorization are unchanged. The header export icon/admin badge/department-button cleanup remains local.

Utilisateurs table actions (local): PersonnelPage owns the export/history icon group, passed through PersonnelTable.headerActions to the right of its sheet tabs; empty/loading/error tables keep the group accessible. PersonnelToolbar now contains only page filters and registration controls. Existing export format permissions and snapshot logic are unchanged.

Readiness indicators (local): dashboard/ReadinessStatus.tsx now uses a neutral compact info/count badge with a hover explanation instead of prominent warning text. TeamPage places the shared badge beside the name; zero-issue records remain unmarked and clicking still opens the same issue details. Dashboard/account consumers share the restrained appearance.

2026-10-30 UI-only release: `rh-compact-ui-20261030T140103Z`; details in [release verification](releases/compact-ui-20261030.md). `FrenchDateField` retains picker drafts separately from serialized form values; Administration uses `AccountList` and guarded detail dialogs.

NIR display exclusivity: `personnel/exclusiveDisclosure.ts` coordinates table, detail and editor reveal requests. A new claim synchronously hides/aborts the previous view; ownership and generation checks reject late responses. Only one plaintext reveal per application window; no values enter the coordinator or query cache.

Account bindings: direct managers are optional in account creation and organization updates. `work_center.person_readiness` reports broken assigned managers, not an absent optional binding; Super Admin is exempt from employee/contract/register completeness checks. New Super Admin UI registration uses the account-only route and no employee record; existing historical employee/contract data is retained.

Super Admin self-demotion (local): organization updates retain the users-wide write lock and last-active-super check but allow an existing Super Admin to relinquish own system rights. Role changes atomically revoke sessions/recovery tokens and audit; UserEditor unlocks own system checkboxes and explicitly clears remembered login before triggering session expiry on any successful own-role change. No actual production roles were changed.

Account detail save (local): AdminPage keeps the Enregistrer submit button in its fixed DialogActions footer. Its native form attribute targets UserEditor.formId inside the organization accordion (mounted while collapsed). The existing validation, role confirmation and self-role signout flow are reused; the duplicate accordion save button is removed. No new mutation API.

Published: account binding, Super Admin self-demotion and detail footer changes are live in `rh-role-handover-20261030T150210Z`; see [verification](releases/role-handover-20261030.md). Both API entry points enforce the same authorization rules.

Mailbox identity (published `rh-email-identity-20261030T155220Z`): `backend/account_email.py` owns live-login normalized mailbox uniqueness and linked register/profile email synchronization. `personnel_register.py` adds Admin-only `/personnel-register/email-availability`; `PersonnelEditor.tsx` prechecks on blur. Deleted logins release their address while keeping historical identity; login, recovery, signature identity and legacy fiche matching avoid claiming retired identities by reused email. See [邮箱规则与部署注意](account-email.md) and [verification](releases/email-identity-20261030.md).

Production consolidation 2026-10-30: only `presence-react-api` (`backend.rebuild_server:app`, 8015) runs for RH. It owns contract reminders, signed PDF retrieval and snapshots. Old units are masked and runtime directories archived; the previous API hostname proxies to 8015 only for link/callback compatibility. Future deployments must not start or update a separate flat API or Streamlit instance. Keep `/etc/presence-app` and `/var/lib/presence-app` as shared current configuration/data paths. See [retirement verification](releases/retire-legacy-20261030.md).

Code/test cleanup (published `rh-leave-dot-20261030T164037Z`): backend modules now use package imports only; retired flat-import branches and their single obsolete import test are archived under `archive/flat-api/`. Unused `fiche_record_belongs_to_user` and no-op contract sync were removed, while historical data migration stays active. `test_support/sqlite_fixtures.py` provides fast disposable SQL fixtures; `scripts/run_tests.py` supports full/targeted regression with slow-case timing. See [test workflow and before/after evidence](test-workflow.md).

Leave notification dot (published `rh-leave-dot-20261030T164037Z`): `useLeaveReviews.ts` shares the existing authenticated pending-team query between `AppWorkspace.tsx` and `LeaveWorkspace.tsx`. The shell polls every 30 seconds only for reviewers; existing leave mutation invalidations refresh it immediately. Pending requests/cancellation requests show a red dot on the collapsed mobile menu button and on the Congés payés navigation label; opening the drawer hides the menu-button dot. Clicking the notified leave entry opens the team's pending filter. No unread state, new API or browser notification permission is introduced. Frontend tests/lint/build passed; isolated 390px synthetic UI verification checked both placements, pending navigation and removal when pending reaches zero. QA harness removed, no real requests modified.

Published verification: [leave dot and package cleanup](releases/leave-dot-20261030.md). The sole API remains on 8015; retired units/ports remain disabled.


## Storage capacity panel (local, 2026-11-04)

`backend/system_storage.py` registers read-only `GET /system/storage` in both package entry points. It rechecks active account/session and Admin/SuperAdmin roles before measuring application/upload/signing filesystems with disk_usage, deduplicates shared devices, hides filesystem paths and reports partial failures. Thresholds: 80% warning; 90% or less than 5 GiB available critical. It does not scan file contents, change data, or measure a separate database server. Registry snapshots remain encrypted database blobs, not a separate directory.

`frontend/src/admin/StoragePanel.tsx` adds a compact storage button below Administration in Espace Admin, with a detail dialog, minute refresh while active, manual refresh and local-machine labeling. No backup, deletion or email automation is introduced. Tests: `test_system_storage.py` covers thresholds, duplicate filesystems, measurement failure, live-role authorization and session revocation.

Published 2026-11-04: storage capacity panel is live in `rh-storage-20261104T093711Z`; production uses Stockage serveur and real server filesystems, local copies retain Stockage local. See [release verification](releases/storage-20261104.md).


## Completion navigation and quieter help (local, 2026-11-04)

`WorkDestination.completion` carries a readiness issue code while preserving exact username/employee identity and the existing draft navigation guard. `ReadinessStatus` links each issue to its completion action. `admin/UserEditor` expands organization settings for manager issues; `AccountPersonnelSection` sends linked editable personnel to the exact detail editor, or opens association search when unlinked. `PersonnelDetail` opens the first missing basic field or entry-period editor. Attendance contract issues open the existing month-specific completion dialog without creating/saving any record.

`completionTarget.ts` maps issue sections and missing basic fields. `components/useCompletionFocus.ts` waits for visible scoped form controls, scrolls/focuses and briefly outlines them, cancelling on unmount and never revealing sensitive numbers. `FrenchDateField` accepts an optional completion wrapper without changing date serialization. No backend permissions, validation, mutation APIs or data are changed. Missing-field rules remain server-owned.

Routine explanation text in account, register, association, history, employment and profile settings is shortened/removed or moved into hover titles. Required labels, errors, local mail-disabled notice, access protections and destructive confirmation remain. Readiness dialogs omit redundant status chips; register detail uses a compact Compléter action.

Validation: 69 frontend tests (including completion routing/missing field selection), lint and production build; disposable browser fixtures verified manager selector, missing first name, email, date of birth, entry date, contract start and unlinked association search focus. No real data edited or messages sent; fixture files/tab removed. Not yet deployed.

Published 2026-11-04: completion navigation and concise help are live in `rh-completion-20261104T104651Z`. All packaged backend files match the preceding storage release; only frontend behavior changes. [Verification and rollback](releases/completion-20261104.md).


## Editable attendance contract refresh (local, 2026-11-04)

`contract_periods.contract_values` shares contract/range metadata construction with new sheets. `backend/attendance_contract_sync.py` compares stable account+contract+month bindings to current contracts. Archived sheets/contracts, nonintersecting/missing contracts, unbound legacy records and signed/in-flight production months retain their exact snapshots. Unsigned active sheets keep their IDs, calendar overrides, exceptions, schedule and other fields; prior metadata is appended to `contract_sync_history` with actor/time.

`GET /attendance/workspace` only reports a read-only refresh-needed flag. The frontend completes `POST /attendance/refresh-contracts` with the observed revision before showing the workspace. The existing checked writer rechecks live permissions, optimistic revision and current contracts/signatures under its users/Presence locks; refresh and leave-ledger writes share the same transaction. Every checked attendance save/open also canonicalizes editable contract metadata, preventing old clients from restoring stale dates. No GET writes or schema migration. `leave_attendance.sync_new_sheets` additionally applies approved leave only to newly included half-days, retaining restoration history and rejecting status conflicts atomically.

`attendance/contractRefresh.ts` awaits canonical config/revision and never retries conflicts. `AttendancePage` adopts refetched revisions only when no dirty, busy or pending calendar work exists; drafts otherwise retain the older revision and receive the existing conflict guard.

Validation: 587 backend tests including real loopback MySQL temporary tables, plus one added end-to-end MySQL scenario for the reported old/new dates with a pure Responsable (refresh, attendance retention, stale-save rejection and signed-source preservation); 73 frontend tests, lint and build. No production data changed.


## Direct-report batch editing (local, 2026-11-04; not deployed)

`frontend/src/admin/ReportsEditor.tsx` adds search, multi-selection and a transfer/removal confirmation to account details. `reportChanges.ts` submits only changed accounts with their original manager for conflict detection. Removing a report may leave it without a direct manager.

`backend/management_reports.py` registers `PUT /users/{manager_username}/reports`: live Admin/Super Admin authorization, ordered user locks, optimistic assignment checks, cycle validation and atomic audit inserts. `main.py` initializes additive table `management_assignment_events`. Only `users.manager_username` changes; department membership/authority, contracts and historical signature identities remain unchanged. Tests: `test_management_reports.py`, the real MySQL temporary-table report workflow, and `frontend/test/reportChanges.test.mjs`.

Account details layout (local, 2026-11-04): `admin/UserEditor.tsx` uses horizontal Informations / Organisation / Contrats et historique / Personnes sous responsabilite tabs. Account access controls and access history are included below the Organisation form; there is no separate access tab. Completion dots sit inline to the right of tab text. Panels remain mounted while hidden to preserve drafts and the footer save form. Completion navigation selects the relevant tab; system-only Super Admin accounts omit personnel and contract tabs.

Account activation layout (local, 2026-11-04): `UserEditor.tsx` supplies a header slot beside the account name to `AccountAccess.tsx`. The existing status and activation controls render there using a portal, retaining one confirmation/state owner. Deletion controls and access history remain under Organisation. Authorization and activation API behavior are unchanged.

Published 2026-11-04: `rh-reports-20261104T140753Z` includes batched direct-report editing, horizontal account details with compact controls, and protected editable contract-snapshot refresh. Earlier local-only notes for these changes are superseded. See [release verification](releases/reports-20261104.md).

Frontend-only publication 2026-11-04: `rh-report-align-20261104T142250Z` serves aligned direct-report name/department rows. Sole API/current remains `rh-reports-20261104T140753Z`; Nginx static root differs intentionally. [Verification](releases/report-align-20261104.md).

Local leave-mail development (not deployed): `backend/leave_notifications.py` owns the transactional `leave_notification_outbox`, recipient resolution and SMTP worker. `leave_requests.py` queues submission/cancellation events; `main.py` registers migration/worker. `frontend/src/leaveLink.ts` and `AppWorkspace.tsx` open the linked request after login. No Admin popup; local/candidate workers must be disabled. See [mail workflow and limitations](leave-notifications.md).

Local extension (not deployed): the leave notification outbox also handles approved/rejected leave and cancellation results, sending to the employee with decision-event comments and a details link. Withdrawal stays silent. No new table columns or frontend changes are needed; see `test_leave_notifications.py` and the MySQL notification workflow.

Published 2026-11-04: `rh-mailnotify-20261104T162306Z` enables new-request and decision-result email notifications, login-protected leave detail links, and a transactional mail outbox. Withdrawal stays silent. Previous local-only notes for this feature are superseded. Sole API and frontend now share this release. [Verification](releases/mailnotify-20261104.md).

Local help guide (not deployed): `frontend/src/help/HelpDialog.tsx` provides the sidebar-bottom question-mark dialog. `helpTopics.ts` holds French step-by-step instructions and filters topics through the shared business/system role policy. The guide opens on the current module; it does not navigate, submit or unmount page drafts. Desktop and mobile drawer share the same entry/state. `helpTopics.test.mjs` verifies role-specific module availability.

Published frontend 2026-11-05: `rh-help-guide-20261105T101144Z` adds the French question-mark guide with left-side topic tabs/right-side steps. Previous local-only help notes are superseded. Nginx static root uses this frontend; sole API/current remains `rh-mailnotify-20261104T162306Z`. [Verification](releases/help-guide-20261105.md).


## Visites médicales (local development, 2026-11-05)

Independent menu alongside Documents: `frontend/src/medical/MedicalVisitsPage.tsx` and `medical/api.ts`. Backend `medical_visits.py` owns employee-bound date/history/private-PDF routes and additive tables; `medical_crypto.py` protects files with a separate AES-256-GCM key ring. Registered by local/rebuild entries after personnel schema migration. Employee: own linked dossier; Responsable: current team read-only via main.user_can_manage; Admin: visible personnel. Owner/Admin corrections retain prior versions; Admin can archive/restore and only Super Admin can permanently delete archived visits. Register deletion retains dossiers referenced by medical visits. See [medical-visits.md](medical-visits.md). Not deployed.

Medical year/optional-month filters use actual visit dates; history groups by month. `/medical-visits/export` builds an audited owner-scoped/Admin-scoped ZIP across all matching pages, organized as year/month/employee. Shared archive filename sanitization and 100 MiB limit come from `document_downloads.py`; at most 100 PDF entries, with atomic failure. Employee registration choices are not restricted by the selected period.

Medical workflow extension (local only): `medical_access.py` reuses the existing team policy for current and historical certificate reads; `medical_history.py` owns versioned correction/history, Admin archive/restore and Super Admin permanent deletion. `medical_visit_versions` and the is_archived column are additive; old encrypted PDFs remain until confirmed permanent deletion. Frontend `medical/MedicalHistory.tsx` shows corrections, archives and role-specific controls. Medical exports return ordinary ZIP files containing unchanged originals under YYYY/MM/Name_employeeID_date_visitID.pdf, without index PDF or links. Per-person downloads include all years and archived visits; the toolbar exports the current filtered scope across pages. Selected IDs are fully authorized. `MedicalExport.tsx` shares download controls between toolbar, employee rows and selected history rows.


Published 2026-11-05: `rh-medical-20261105T161729Z` adds Visites médicales, versioned encrypted certificates, team/owner permissions, Admin archives and Super Admin deletion, and ordinary year/month ZIP downloads (filtered, individual full-history or selected). Previous local-only medical notes are superseded. Sole API/current and frontend share this release; original 43 business tables preserved, three medical tables initialized empty. Independent server-generated medical key and guarded single-service rollback backed up. See [medical release verification](../doc/releases/medical-visits-20261105.md).


## Annual approved leave duration (local, 2026-11-06; not deployed)

`backend/leave_totals.py` computes read-only calendar-year totals for the existing authorized `/person-profiles/{username}` response. `team/PersonProfile.tsx` displays an annual total, and `ProfileMonth.tsx` adds monthly days. It reads approval-time `leave_attendance_cells` snapshots, counts each date/half-day once across workspace copies and overlapping approved requests, and excludes rest/holiday cells. Archived approved requests still count; cancellation removes time only after approval. Future approved dates count. Requests without verifiable synchronization history are marked incomplete, never guessed using today's planning. No migrations or attendance/signature changes. See [leave totals](leave-totals.md).
