# Presence 模块索引与维护说明

2026-10-22 发布回归：新增 `test_account_personnel.py` 的 10 项隔离 HTTP 测试，覆盖本人只读、负责人隔离、关联/替换/解除、唯一占用、档案版本、失效会话、审计失败回滚和敏感字段排除。`test_flat_backend_imports.py` 检查保留 API 的 flat 导入；`team_directory.py`、`attendance_permissions.py`、`account_purge.py` 同时支持包导入和 flat 导入。启用本地连接级 MySQL TEMPORARY 测试后全量 369 项通过；前端 43 项规则测试、TypeScript、lint 和生产构建通过。测试未修改真实员工资料、未发送邮件或创建签署。发布版本 `personnel-profile-20261022T125536Z` 已启用员工名册、Profile 关联、归档权限、姓名展示与页签布局等累计改动；下方“本地、未部署”描述保留为各阶段历史记录。生产 NIR 仍关闭，不导入本地资料，不复制本地 SuperAdmin 赋权。详情与回退边界见 `rebuild-release-readiness.md`。

2026-10-22 个人设置页签布局（本地，未部署）：`profile/ProfilePage.tsx` 将账号设置与员工名册资料改为两个 sheet 页签，切换时保持表单挂载以保留未保存输入；保存工具栏只在账号设置页签显示。`AccountPersonnelSection` 的 `embedded` 模式直接展示资料，其他入口仍保留折叠卡片。接口和权限不变。

2026-10-22 账号 Profile 与员工名册关联（本地，未部署）：`profile/AccountPersonnelSection.tsx`、`accountPersonnelApi.ts` 在 Administration 的账号详情、个人设置和本人/Admin 的月度 Profile 复用名册数据卡。Admin/SuperAdmin 可搜索未归档且未被其他账号占用的登记记录，关联、更换或解除关联；普通账号（含负责人）只能查看自己的关联姓名、国籍、性别、生日、地址、职位、资格、人员类型及任职期间。仅直接负责人查看下属月度概览不获得这些名册资料；`person_profiles` 的 `can_view_personnel_data` 只是界面标记，实际仍由独立后端接口鉴权。

新接口由现有 `personnel_register.register_personnel_routes` 注册并复用 no-store/安全错误处理：`GET /account-personnel/{username}` 仅本人或 Admin；`GET .../choices` 和 `PUT .../link` 仅 Admin。使用现有唯一 `linked_username` 关系，不创建第二套关联，不按姓名/邮箱猜测。写操作复用 users 排序锁、员工版本和当前关联检查，更换时原档案解除关联与新档案关联及两条审计同事务提交；已占用/过期版本/新增关联归档档案被拒绝。社保号、密文、存在标记、凭据和名册修改日志均不进入 Profile 的只读数据；原社保号专用接口权限不变。解除关联不删除员工资料、任职历史、考勤或合同。`useRegistryAction` 同步失效账号资料卡与选择列表，关联后刷新姓名相关缓存。

本轮仅做 Python AST、TypeScript、lint 和构建检查，未新增或运行功能测试、未实际关联/解除任何真实资料，也未部署。账号原有邮箱/密码设置仍可修改，普通员工“只读”指名册资料及关联关系。

2026-10-22 账号列布局与考勤历史入口（本地，未部署）：`admin/AccountOption.tsx` 与 `accountOptionLayout.ts` 为账号、负责人、员工分配共用三列菜单，姓名/邮箱、职位/人员身份、管理权限分别对齐；选中状态只显示姓名及必要权限。`rolePolicy.ts` 将标签统一为 Admin / Super Admin，业务身份与管理权限分开投影，内部权限标识不变。`main.admin_job_titles` 只为 Admin/SuperAdmin 的 `/list-users` 补充明确关联名册的 `job_title`，仅读取关联与职位列，普通员工/经理不新增此字段，也不扩大公共目录/会话字段。

`attendance/types.accountSheetPeriods` 仅从已授权配置的明确账号关联及已存月份生成历史入口，区分普通/归档；切换不回写考勤或合同。考勤页面提示该人员在当前月份无表，并提供跳转已有月份按钮；选中当前月已有人员时同步激活其页签。`attendance/open` 在无当前有效合同但已有明确绑定同月未归档快照时允许打开，原合同快照/考勤数据保持原样；新建仍要求有效合同，已归档考勤仍须先恢复。名册任职日期和考勤合同继续独立维护，本轮未改用户的实际合同日期。

本轮验证：新增 9 项历史考勤打开回归、8 项 Admin 职位字段隔离回归，完整后端 358 项中 352 通过、6 项 opt-in MySQL 跳过；前端 43 项通过、全量 lint / TypeScript / 构建通过。真实本地 GET 核对职位仅按显式关联投影、2026 年 9 月考勤仍在，没有创建考勤、修改合同、发送邮件或部署；浏览器布局仍待人工查看。

2026-10-22 人员名称显示（本地，未部署）：界面统一采用明确关联员工名册的 `full_name`，未填写/未关联时使用有效邮箱 `@` 前缀；两者缺失时显示中性提示，不以 `username` 兜底。`backend/person_names.py` 仅批量读取名册姓名/关联列及必要账号邮箱，不读取社保号、生日、地址或凭据；旧服务器缺名册表时安全降级。`frontend/src/personDisplayName.ts` 统一界面优先级。账号 ID、授权、路由、分组关系与历史存储不变，不按相似姓名或邮箱猜测关联。

姓名投影覆盖登录会话、账号/负责人选择、组织架构、Documents 员工/上传人/归档人、考勤页签、签署预览、Dashboard、Profile、请假和历史操作者。`attendance/workspace` 的 `person_labels` 为只读显示元数据，刷新姓名不改变未保存考勤或历史导出姓名；签署人的名称使用当前已授权账号的关联姓名。名册更正姓名或账号关联成功后失效相关显示缓存。账号删除确认优先输入当前邮箱，接口兼容原用户名确认；账号设置中的明确技术凭据字段仍可编辑。

2026-10-22 姓名功能回归（本地，未部署）：全量 Python discover 共 341 项，335 通过、6 项 opt-in 跳过；这 6 项本地 MySQL TEMPORARY 表测试另行全部通过。新增 `test_person_names.py`（27 项）与 `test_person_name_workflows.py`（18 项），并补账号删除邮箱确认测试，覆盖同名身份隔离、缺名册兼容、历史快照/草稿不改写、权限及敏感字段隔离。前端 `npm test` 共 35 项通过，TypeScript、全量 lint 和构建通过。真实本地 API 9 个只读接口返回 200，4 个无登录请求及篡改/旧版会话被拒绝；没有修改真实员工、上传文件、发送邮件或签署，浏览器交互仍需人工验收。

本轮修复异常邮箱 fallback 校验前后端不一致；`personDisplayName.ts` 的 `personOptionLabel` 为下拉框提供姓名及邮箱辅助识别同名人员；`admin/accountIdentity.ts` 将删除确认身份与人名标签分开。`attendance/types.ts` 的 `attendancePersonName` / `withAttendanceLabels` 只更新展示元数据，不改配置和导出快照。`leave_requests.historical_actor_name` 在当前账号无法解析时采用既有可信姓名快照，拒绝原始账号 ID 和异常邮箱，不推断或改写审计身份；没有姓名快照的账号/名册操作仍显示中性旧账号提示。

2026-10-22 本地权限验收：`ChristopheH` 已通过显式 local-only 脚本赋予 SuperAdmin，保留员工显示；正式隐藏尚未实现，也未部署。赋权前仅角色元数据备份在被忽略的 `.local-test/rebuild/role-backups/`，角色变化递增会话版本，需重新登录；回退角色也应递增版本，不可恢复旧会话版本。当前后端回归 294 项中 288 通过、6 项 opt-in 跳过，这 6 项真实本地 MySQL TEMPORARY 测试已另行全部通过，无真实记录删除、邮件或签署。前端 7 项既有测试、完整 lint、TypeScript 及构建均通过；浏览器交互仍需手动验收。

权限与资源维护入口：`backend/role_policy.py` 提供角色继承、权限变更/最后 Super 保护、实时身份重查、角色审计和不可复用用户名；`account_purge.py` 只移除已停用且业务依赖已处理的账号，不级联删除别人的资料。`resource_maintenance.py` 由两个 React 后端入口显式注册，提供部门/请假/年访谈的分页归档与受控永久删除；`destructive_audit.py` 的独立最小删除/归档审计只接受类型、ID 和数量，不接收自由文本、NIR 或密文。维护工具不提供删除这些操作审计的接口。`frontend/src/rolePolicy.ts` 统一判断最高角色并去除冗余显示；`admin/AdvancedMaintenance.tsx` 提供紧凑管理窗口，`components/PermanentDeleteDialog.tsx` 共用精确编号确认。

考勤保存由 `attendance_permissions.py` 核对已有 ID：普通 Admin 不得通过整份配置省略来删除；Super 也须先归档并在 `/attendance/save` 明确提交 `deleted_ids`，与被移除 ID 完全相同。`is_archived` 只改变展示/使用状态，保留原月份绑定和合同快照。已归档合同不用于新建考勤，历史快照仍可导出；已归档考勤不参与新签署。`scripts/promote_local_superadmin.py` 仅可显式操作 127.0.0.1:3307 的 `presence_local_rebuild`，核对已存在账号及邮箱，不创建员工、不设置密码、不在启动时自动赋权。

2026-10-22 资源权限升级（本地，未部署）：SuperAdmin 继承 Admin，Admin 保留查看/创建/编辑/归档但不能永久删除；本地 ChristopheH 的测试身份暂时可见，正式隐藏以后另做。名册永久删除仅 SuperAdmin，须已归档、版本及登记号/姓名二次确认，清除该名册员工、任职和个人资料历史，同时保留 `permanent_deletion_events` 独立最小元数据审计，不记录 NIR/密文。Documents 永久删除仅 SuperAdmin 且先归档、确认文件 ID，员工不再自删；普通员工本人/上传者可归档恢复，非 Admin 经理仍只读。合同增加 `is_archived`，GET 保留全历史，Admin 归档/恢复，SuperAdmin 删除须归档和确认，任意保留考勤快照引用时拒绝。写操作复用 `role_policy.checked_resource_actor`，排序锁用户并重查状态/会话及角色；删除审计与数据库变更同事务，文件清理失败返回 `cleanup_pending`。仅隔离测试，未实际删除真实记录。

2026-10-22 工作区布局（本地，未部署）：全部八个主页面的标题移到应用顶部栏，`workspacePages.ts` 集中维护法语标题和悬停说明；标题停留约 650ms、键盘聚焦或触屏长按均可查看说明。`WorkspaceToolbar` 只挂载当前页操作，月份、选人、注册、刷新等从重复页面头迁出；正文减小边距并使用可用宽度，错误、未保存状态及邮件/签署确认继续显示。考勤保存仍固定右下角；账户设置的顶部提交按钮通过 HTML form 属性关联原表单，保留原生校验。未部署、未发送邮件或修改业务数据。

仪表板后续紧凑调整：`WorkspaceToolbar` 支持 `placement="inline"`，通过独立宿主把刷新放在标题同行，不占第二行工具区；其他页面默认工具栏布局不变。桌面端人员分布与统计卡片并排，环形图 190px、图例数字固定列对齐，较窄屏自动上下排列；缩小内边距、悬停明细占位和折叠栏高度。统计与刷新权限、请求、确认逻辑不变。

2026-10-22 员工名册第一版（本地实现，未上线）：`frontend/src/personnel/` 与 `backend/personnel_register.py` 提供独立的 `Registre du personnel`。人员身份与登录账号分开，姓名唯一必填，账号选填一对一关联，任职经历可追加、更正并保留完整修改历史。社保号已从早期预留升级为本地受控加密录入/查看，当前仅使用合成号码测试；生产默认关闭，尚未部署。主列表每人一行，详情查看多次入职；后台分页、搜索、归档和人员类型/状态筛选。不从现有账号生成员工，普通 Admin 用归档替代永久删除；SuperAdmin 受控永久删除见本页权限升级说明，仍不提供 Excel 导出。资源权限升级前的验证基线：NIR 专项 52 项通过（29 单测、18 SQLite HTTP、5 真实 MySQL TEMPORARY）；全量 Python 237 项中 231 通过、6 项 opt-in 跳过，新增 MySQL 5 项已单独通过。前端 TypeScript/lint/build、7 项既有 draftRules 测试通过；浏览器视觉/输入交互复核受 URL 策略限制，仍需用户本地确认。角色、字段、存储、接口及验收边界见 [员工名册第一版](personnel-register.md)。

NIR 配置与边界：`backend/personnel_sensitive.py` 使用 AES-256-GCM 和 `PRESENCE_PERSONNEL_NIR_KEYS_FILE` 指向的独立 keyring，生产必须显式设置 `PRESENCE_PERSONNEL_NIR_ENABLED=1` 且密钥有效才开启。本地入口自动 provision 被忽略的 `.local-test/rebuild/keys/personnel-nir.json`，不生成生产密钥；POSIX 文件创建为 0600，Windows fresh 环境须单独配置 NTFS ACL，不能靠 mode 参数实现。密钥单独备份，rotation 保留旧 key。公开 DTO 仅 `has_social_security_number`，输入未传保留、null/空白清除、非空替换；专用 reveal 先提交查看审计、HTTP no-store，前端显式显示后 30 秒或离开即清除且不进查询/浏览器存储。审计仅记录是否存在及 changed/viewed 事实，不留原号/密文；格式规则不核验真实身份或 checksum，其他个人资料、自由文本及附件不属于此字段加密范围。

2026-10-21 登录标识更新（独立补丁已上线 `email-login-20261021T151000Z`）：`backend/main.py` 的 `/login` 保留 username 请求字段，接受已绑定邮箱或用户名，忽略大小写及标识两端空格；先唯一解析为实际 username，再签发会话和记住登录令牌。用户名与邮箱跨账号冲突或历史重复邮箱导致多匹配时统一拒绝，不猜测身份；密码不去空格。本地保留停用/删除登录与会话版本校验。`components/Login.tsx` 提示邮箱或用户名，不新增账户、不改原密码或邮箱，不代表已完成邮箱所有权验证。线上按两套实际源码只替换 login 标识解析，保留各自既有认证逻辑；前端从已上线的 4d6e0ad 基线只更新登录框，不包含本地档案、停用或删除登录等新功能。185 项后端回归、7 项前端测试、构建与当前源码 lint 通过；两套实际 API 补丁另有 10 项隔离登录测试通过，MySQL 临时表补充真实 bcrypt 邮箱登录通过。线上 7 个现有身份的只读映射和会话检查通过，无真实密码修改、邮件或签署；备份及回退见 doc/rebuild-release-readiness.md。

2026-10-15：React 重制版已发布至 `https://app.presence.217.65.145.103.nip.io/`，版本 `react-20261015T101334Z`。新版服务 `presence-react-api` 使用 127.0.0.1:8015；原 Streamlit 与旧 API 保留运行以供回退。服务器使用原线上数据库与文件，未导入本地快照。发布与回退记录见 `rebuild-release-readiness.md`。

2026-10-16：Documents 搜索及三档卡片密度已随前端 `documents-search-density-20261016T091655Z` 上线；回退记录见 `rebuild-release-readiness.md`。

本索引用于按功能定位代码，减少每次开发重复阅读整个项目。当前前端入口是 `frontend/src/main.tsx`，线上后端入口是 `backend/rebuild_server.py`（复用 `backend/main.py`）。旧版 UI 仅供回退，职责映射见 `architecture.md` 的旧版表。路径均相对于仓库根目录。

Documents 的功能、使用流程和上线状态见 [Documents 开发报告](documents-report.md)。
总入口、脚本依赖和 Mon équipe 阅读顺序见 [脚本职责与应用入口](architecture.md)。
2026-10-14：Mon équipe 只读目录与响应式架构已上线，完整回归 78 项通过；发布后 7 个现有账户只读接口检查通过。细节见 `refactoring.md`。
后续拆分、重复调用与状态风险见 [代码优化与维护计划](code-audit.md)。该计划中的建议尚未实施。

新技术栈试点：[Congés payés 申请模块](leave-pilot.md)。React 入口位于 `frontend/`，独立本地预览保留；2026-10-15 已随重制版发布，旧版 Streamlit 保留供回退。
2026-10-14 本地重构第一阶段：[独立真实数据测试版](local-rebuild.md)。默认新前端改为连接 `backend/local_rebuild.py`（8013）及独立 MySQL（3307），不再选择虚构演示身份；优先迁移考勤编辑、单份及批量文件生成，尚未发布。
新前端主题：`frontend/src/ThemeModeProvider.tsx` 管理 Light/Dark 和浏览器偏好持久化，`themeMode.ts` 提供共享上下文，`theme.ts` 提供两套配色，`components/ThemeToggle.tsx` 是登录及工作区共用切换按钮。
2026-10-15：考勤表内三个区域改为横向页签（日历、个人与合同、每周安排；实习生第三页为报酬与出勤），默认显示日历，窄屏可横向滚动。隐藏面板保留状态，日历只在激活时请求；个人字段与安排仍由 `FicheEditor.tsx` 按 section 渲染，切页不写入服务器。
新版日期输入统一由 `components/FrenchDateField.tsx` 使用 MUI DatePicker 与 Day.js；`main.tsx` 设置法语本地化，显示 DD/MM/YYYY（占位 JJ/MM/AAAA），接口继续使用 ISO 日期。应用于考勤合同、Admin 合同、面谈、请假；保留必填、起止限制和禁用权限，非法日期不能保存考勤。
新版各页面可折叠区统一默认收起：`theme.ts` 设置 MuiAccordion 默认值，页面不再单独默认展开；考勤日历受控状态也从 false 开始。用户主动展开后，当前会话切页继续保留页面状态，不强制关闭未保存表单。
新版视觉公共组件：`components/DepartmentBadge.tsx` 提供固定列宽、长名称省略与全名提示；`departmentColor.ts` 共享稳定部门配色（标签与人员分布图）；`attendance/AccountOption.tsx` 提供部门/姓名分列选人；同一标签共用于团队名单与 Admin 部门选择。`theme.ts` 统一控件高度、表格及折叠区；`index.css` 管理页面间距、人员控件与窄屏月份对齐。
新版配色在 2026-10-15 按用户反馈改为蓝色系：深色模式主色 `#7aa7f8`、浅色模式 `#285fbd`；保留深色背景 `#0e1117` 与灰色控件 `#262730`。姓名高亮、常规按钮、标签页与日历选择框使用主题主色；错误和删除继续采用语义红色。
新版日历拖选由 `frontend/src/attendance/useCalendarSelection.ts` 管理：鼠标/触摸 Pointer Events 沿经过的半天格添加或取消选择，支持键盘点击；`AttendanceCalendar.tsx` 仍负责显示和统一应用状态。
新版日历状态颜色集中在 `frontend/src/attendance/calendarStatus.ts`，共用于半天格、状态选择器及图例，分别适配 Light/Dark；拖选使用主题色边框保留状态底色。成功提示统一由 `theme.ts` 使用中性灰色呈现。

## 页面与共用模块

| 功能 | 入口 | 共用模块与职责 |
| --- | --- | --- |
| 员工名册（本地） | `frontend/src/personnel/PersonnelPage.tsx` | 独立 Admin 菜单，注册员工与登录账号不同；精简响应式表格、行尾「…」完整档案、`PersonnelToolbar.tsx` 顶部注册/搜索/筛选/分页；人员编辑、任职经历、账号关联、合同预填及修改历史；`SensitiveNumber.tsx` 专用短时查看、`useSensitiveNumberCapability.ts` 仅缓存能力标记；复用草稿保护、统一请求与法语日期。后端规则/存储/加密/路由分别在 personnel_rules/store/sensitive/register，入口显式注册，不改旧合同/考勤 |
| 应用栏页面操作槽 | `frontend/src/workspaceToolbar.ts`、`frontend/src/components/WorkspaceToolbar.tsx`、`frontend/src/workspacePages.ts` | AppWorkspace 提供吸顶宿主、逐页可见状态与高度测量；统一 h1 标题/悬停说明由 workspacePages 管理，八个页面通过 Portal 提交顶部工具栏，业务状态留在页面，隐藏页不渲染操作；名册按剩余可视高度滚动，其他正文自然纵向滚动 |
| 考勤校验与生成 | `fiche_exports.py` | `fiche_name`、`validate_fiche`、`generate_fiche_file`，供单份、批量导出和批量签署共用 |
| 考勤日期序列化 | `fiche_serialization.py` | 保存时日期转 ISO 字符串；读取时恢复日期对象 |
| 默认工作安排 | `planning.py` | 工作日名称、默认时段及新记录的默认安排；不覆盖已保存安排 |
| 日历显示与修改 | `calendar_view.py` | 半天状态解析、例外记录同步、批量编辑与预览 |
| 日历浏览器组件 | `calendar_selector_component/index.html` | 日历交互与拖动选择；组件加载由 `calendar_view.py` 管理 |
| 月份选择 | `period_selector.py` | 法语月份名和同一页面月份选择器之间的状态联动 |
| 合同与月份匹配 | `contract_periods.py` | 合同快照、月份内有效区间、跨月例外过滤及合同文件名 |
| Excel 模板生成 | `ExcelGen.py` | 将工资与考勤数据写入 `Fiche_Exemple.xlsx` |
| 实习生 Word 生成 | `DocxGen.py` | 使用 `template_stagiaire.docx` 生成文件 |
| Documents ZIP 内容生成 | `document_downloads.py` | 逐个鉴权下载、文件名清理、100 MB 上限；失败时不返回部分 ZIP |
| 新版带薪假申请 | `frontend/src/App.tsx` | React + TypeScript + Vite + Material UI；个人/团队申请、分页、过滤、系统内提醒、申请/审批窗口及操作历史，接口集中在 `src/api.ts` |
| 新版总入口及请假页面 | `frontend/src/AppWorkspace.tsx`、`frontend/src/LeaveWorkspace.tsx` | `App.tsx` 只管理会话，外壳提供统一导航，请假页面只保留业务内容，导航统一由 AppWorkspace 提供；界面主题与角色菜单共用 |
| 新版页面加载和草稿提醒 | `frontend/src/AppWorkspace.tsx` | 各模块独立加载与失败重试，已打开模块保留状态；`draftGuard.ts`、`draftRules.ts`、`components/DraftGuardProvider.tsx` 共用草稿注册、确认与刷新/关闭提醒；上下文提供不含字段值的待保存模块标签，外壳持续显示法语提示，考勤保存区区分待保存/干净/待应用日历状态；未应用的日历选区在切换员工、打开其他档案或切月前确认，覆盖考勤、管理表单、Documents 上传、请假、面谈、账户设置；提交期间阻止工作区切页/退出 |
| 新版考勤页面 | `frontend/src/attendance/AttendancePage.tsx` | 保留完整 JSON 草稿、稳定 ID 分页、保存、导出与勾选；`AttendanceToolbar.tsx` 顶栏月份/选人/打开，`PeriodToolbar.tsx` 共用紧凑响应式月份控件；日历 showPeriodSelector 保留默认兼容，本页关闭重复选择器；保存固定在当前考勤页面右下角，`components/BatchExportResult.tsx` 共用完整导出结果、期间选区快照、失败原因与 JSON 明细（失败不推断单项成功）；单份 Excel/Word 和 PDF 收进 `FicheExportMenu.tsx` 的 More 菜单；编辑、日历及月份拆至同目录组件，API 复用总请求工具 |
| 新版 Documents | `frontend/src/documents/DocumentsPage.tsx` | 单一吸顶月份栏同时控制上传与文件列表，保留草稿切月确认并清空旧月份勾选；待处理/归档、类型筛选、勾选及批量下载；`documentSearch.ts` 在已鉴权列表内按文件名、编号、员工、上传人及类型说明进行多词搜索，忽略大小写和重音；搜索与月份/文件夹/类型组合，改变搜索词清空勾选，不扩大后端权限；列表提供 Compact/Standard/Confort 三档卡片大小，以放大/缩小图标调整响应式列数、间距和标题字号；偏好仅保存在浏览器 localStorage，默认 Standard，不改变勾选或文件；操作按钮保留 44px 点击面积。`DocumentUpload.tsx` 上传弹窗（本地改版：日期栏按钮打开，显示绑定月份，成功关闭，取消/Esc/遮罩关闭检查本弹窗草稿，提交时禁止关闭，错误在弹窗内显示），`DocumentCard.tsx` 卡片及操作确认（视觉版：预览/下载保留常驻图标，改月/归档/删除收入更多菜单；删除保留确认，统一 44px 点击面积、法语提示及 aria-label），`DocumentPreview.tsx` 通过原鉴权下载接口获取 Blob，在响应式弹窗中展示图片（内存 data URL）或延迟加载 `PdfPreview.tsx` 用 PDF.js/Canvas 渲染 PDF，支持翻页及缩放；Vite 将版本化字体、CMap 和解码器复制到被忽略的 public/pdfjs 并打包，全部同站加载，保持现有 CSP 不变；提供新标签页/下载回退、取消加载与关闭后释放 URL，不持久化或发送文件至第三方；`api.ts` 类型与请求；共用 `src/download.ts` 下载工具 |
| 账号停用与会话撤销（本地） | `backend/account_lifecycle.py`、`frontend/src/admin/AccountAccess.tsx` | additive users.is_active/session_version 与 account_status_events；Admin 锁定后重新验权，不可停用本人；PATCH /users/{username}/activation，GET /users/{username}/activation-history；撤销会话、记住登录及重置链接，重新启用需重新登录；旧 DELETE /delete-user 返回 409，不删历史数据。组织目录/人数统计排除停用账号，Admin 保留查询和重新启用；新签署拒绝停用的签署双方 |
| 新版个人月度档案（本地） | `frontend/src/team/PersonProfile.tsx`、`backend/person_profiles.py` | `GET /person-profiles/{username}?year=YYYY`；每次从数据库重新核对本人/当前直接负责人/Admin，授权后聚合全年绑定考勤、文件归档数、跨月请假及可追踪签署发送回执，不返回其他人员配置、文件内容或签署链接；Mon équipe 提供本人按钮、人员姓名和授权范围的档案选择。新版 server/local 入口注册，旧主入口不变 |
| 新版 Mon équipe | `frontend/src/team/TeamPage.tsx` | 以横向标签页切换人员名单和公司架构，默认显示名单；按角色请求本组目录或授权员工列表，支持搜索且切页保留搜索内容；`OrganizationCards.tsx` 在架构标签页直接渲染响应式部门卡片，Équipe 标签前有装饰性人头图标，`api.ts` 定义只读请求与类型；服务端复用 `organization_model.py` 投影 |
| 新版 Administration | `frontend/src/admin/AdminPage.tsx` | Admin 入口、账户搜索/选择、操作状态和查询更新；`CreateAccount.tsx` 新建账户弹窗（顶部按钮打开，成功后关闭，取消/Esc/遮罩关闭检查本窗口草稿，提交时禁止关闭，错误在窗口内显示）与 `UserEditor.tsx` 共用 `OrganizationFields.tsx`；`UserEditor.tsx` 角色/部门/负责人及关联人员，角色变化提交前显示新旧权限确认，`OrganizationFields.tsx` 禁止在编辑本人时取消 Admin；后端组织更新事务串行锁定角色修改并重新读取操作者权限，禁止自行移除 Admin，提交后记录角色变更操作者和前后角色；`ContractEditor.tsx` 合同历史，`GroupEditor.tsx` 部门创建/修改，`api.ts` 请求类型；复用原管理接口，创建功能由 `/ui-capabilities` 控制，本地后端仍禁止新增账户 |
| 新版能力配置 | `frontend/src/capabilities.ts`、`backend/ui_capabilities.py` | `GET /ui-capabilities` 只返回环境与布尔功能开关；界面无法核对时禁用新增账户和邮件，权限仍由各业务接口执行；本地入口另外拦截创建、找回邮件与提醒发送/重发 |
| 新版仪表板 | `frontend/src/dashboard/DashboardPage.tsx` | 合同到期、面谈统计及名单、Admin 提醒偏好及发送/重发确认窗口；`PersonnelDistribution.tsx` 复用已认证组织目录，以轻量 SVG 环形图及可操作图例展示部门人数/比例；悬停、键盘聚焦和触屏选择显示明细，按 username 去重并仅计实际 member 归属，未分配和多部门账号单列，不把管理多个部门重复计算，刷新共用组织缓存；当前无可靠测试账号标记，统计包含目录返回的全部账号（本地试版）。`InterviewEditor.tsx` 编辑面谈，`api.ts` 共用原接口；本地禁用真实提醒发送 |
| 新版账户设置 | `frontend/src/profile/ProfilePage.tsx` | 本人用户名/邮箱/密码修改及当前密码确认；成功后交给 `App.tsx` 更新会话令牌、清理查询缓存并重载账户状态 |
| 新版持久登录与密码重置 | `frontend/src/authSession.ts`、`components/PasswordRecovery.tsx` | 可选记住登录、恢复/撤销；访问令牌保存在 sessionStorage、持久随机令牌保存在 localStorage，密码不持久化；支持 `?token=` 重置入口，本地不发送找回邮件 |
| 新版考勤签署 | `frontend/src/attendance/SignaturePreview.tsx`、`backend/attendance_signatures.py` | 单份/勾选人员预检、确认及逐份发送；直接复用旧版签署服务。`signature_attempts.py` 用本地 SQLite 记录提交回执，防同一次请求重复发送；仅沙盒配置就绪时启用 |
| 签署链接与耗时 | `backend/esign.py`、`frontend/src/attendance/SignaturePreview.tsx` | 沙盒只展示返回的签署页，不展示 PDF/区域设置页；依响应中的邮箱标注签署人，无法核实时保留通用标签；记录生成/PDF/服务商耗时，前端显示等待秒数与明确完成状态；批量预检数量、当前处理人和进度，以及已创建/未发送/待核查结果分开，超时不得算作确定未发送，不自动重试；可下载不含签署链接的 JSON 结果，发送中注册关闭提醒 |
| 本地 PDF 专用公网入口 | `backend/pdf_test_gateway.py`、`scripts/start_pdf_test_tunnel.ps1`、`scripts/stop_pdf_test_tunnel.ps1` | 独立 8014 端口，只读取随机令牌对应且未过期的 PDF；Cloudflare 临时隧道只连接此入口，未挂载 8013 私有 API；启动/停止和临时地址写入本地忽略目录 |
| 指定人员正式邮件测试 | `backend/local_signing_policy.py`、`scripts/stop_live_signing_test.ps1` | 默认沙盒；用户明确授权后，可在本地私有配置限时开放唯一员工/负责人邮箱组合；新适配器与旧发送服务均校验，到期拒绝发送，支持撤销并恢复沙盒 |

## 后端模块

| 模块 | 职责 | 主要接口 |
| --- | --- | --- |
| `backend/main.py` | 应用入口、环境配置、数据库连接、认证、组织权限、用户、提醒、面谈、考勤 JSON 保存和启动迁移 | `/login`、`/restore-session`、`/groups`、`/list-users`、`/dashboard`、`/get-config/{user_id}`、`/save-config/{user_id}` 等 |
| `backend/personnel_register.py`、`personnel_store.py`、`personnel_rules.py` | 本地名册、实时 Admin/SuperAdmin 鉴权、增量密文列、版本冲突、普通及 NIR 元数据审计；Super 删除已归档名册并保留独立最小删除审计；所有名册响应禁止缓存，校验错误不回显原输入 | `/personnel-register`、`/accounts`、`/capabilities`、`/{id}`（含受控 DELETE）、`/periods`、`/archive`、`/history`、`/{id}/social-security-number/reveal` |
| `backend/personnel_sensitive.py` | 独立 keyring、NIR 格式规范化、AES-256-GCM 与员工 ID 绑定、配置/篡改错误失败关闭；本地显式初始化 helper | 由名册路由和本地入口调用；不派生会话密钥、不自动配置生产密钥 |
| `backend/document_numbers.py` | 按巴黎上传日期分配 YYYYMMDD-0001 编号；启动时增量建列、唯一索引和每日计数表，历史文件按 created_at/id 补号；事务行锁保证并发分配，删除不回收编号 | 由 Documents 启动迁移与上传共用 |
| `backend/documents.py` | 文件编号展示/下载前缀、文件上传、访问范围、元数据、月份修正、归档/恢复；永久删除仅 Super 且先归档、确认 ID，提交后文件清理失败显式返回 cleanup_pending；ZIP 逐份复用下载鉴权，最多 100 份、100 MiB，不返回部分 ZIP | `/transport-receipts`、`/transport-receipts/assignees`、`/transport-receipts/{receipt_id}`（受控 DELETE）、`/file`、`/period`、`/archive`、`POST /transport-receipts/export` |
| `backend/esign.py` | LibreOffice 转换、临时 PDF 存储、ClawShow 请求及接口错误转换；注册函数返回原发送服务供 React 适配器复用，原路由不变 | `/fiches/pdf`、`/send-fiche`、`/esign/documents/{document_token}` |
| `backend/signature_flow.py` | 根据员工账户寻找负责人，检查签署邮箱与权限，组装 SIGNER 顺序及员工或实习生的两个签名锚点 | 由 `backend/esign.py` 调用 |
| `backend/signature_placement.py` | 从最终 PDF 查找唯一签名文字，按 ClawShow 当前的 A4 参考坐标换算；拒绝缺失、重复及不支持的旋转或裁剪页面 | 由 `backend/esign.py` 调用 |
| `backend/contract_history.py` | 合同历史表增量迁移，GET 保留归档历史，正常合同日期防重叠，Admin 归档/恢复；Super 确认删除已归档且无任何考勤引用的合同 | `/users/{username}/contracts`、`/users/{username}/contracts/{contract_id}`（含受控 DELETE）、`/{contract_id}/archive` |
| `backend/attendance_permissions.py` | 限制经理保存内容为权限范围内的考勤字段 | 由 `/save-config/{user_id}` 调用 |
| `backend/team_directory.py` | 登录用户查看基础架构；目录响应新增 `organization` 投影供 React 使用，保留旧字段；按服务端部门归属查询员工本组名单，不暴露合同、完整邮箱或财务资料 | `/organization-directory`、`/my-team` |
| `backend/leave_requests.py` | 带薪假接口，注入认证/连接；跟随当前直接负责人，Admin 代审，员工审批前撤回 | `/leave/session`、`/leave-requests`、`/{id}`、`/{id}/decision`、`/{id}/withdraw` |
| `backend/leave_rules.py`、`backend/leave_store.py` | 日期/半天/权限与存储分离；事务内保存状态和操作历史，建表；账号关联字段排序规则按 `users.username` 对齐，保留旧数据；入口账号改名/删除时同步或解绑申请归属 | 正式 MariaDB；独立本地预览使用 SQLite |
| `backend/ui_capabilities.py` | 提供环境及账户创建/邮件开关，不含密码、密钥或数据；不替代鉴权 | `GET /ui-capabilities` |
| `backend/rebuild_server.py`、`server_signing_readiness.py` | 2026-10-15 已上线的新版服务器入口，完整仓库运行；复用原鉴权与保存、添加 React 考勤/签署路由；服务器密钥与公开 PDF 地址就绪检查、持久回执存储 | 新版 `/attendance/*`；部署样例在 `deploy/` |
| `backend/local_rebuild.py` | 独立真实数据副本入口，固定本地连接、独立会话/文件、关闭真实邮件及正式签署 key；不部署到线上 | 复用 `main.py` 路由并挂接新考勤接口 |
| `backend/attendance_api.py` | React 考勤适配；已有人员/合同调用（accounts.group_names 返回有效所属部门名，选人框前缀无分组显示 Unknown）、旧日历规则及原文件生成；`attendance_revision.py` 计算内容版本，`main.save_config_checked` 在事务锁内比较版本，不新增用户 | `/attendance/workspace`、`/attendance/save`、`/attendance/open/{username}`、`/attendance/calendar/{id}`、`/attendance/export` |

Documents、eSign 和合同模块通过注册函数挂到同一个 FastAPI 应用。数据库连接、认证与授权依赖由 `main.py` 显式传入，子模块不反向导入 `main.py`。接口路径继续兼容原前端。

本地可按包导入 `backend.main`；生产可在后端目录使用 `uvicorn main:app`。新增后端模块必须随入口文件一起部署。

## 主要数据关系

- `users`：账户、角色、负责人及兼容的合同摘要字段。
- `organization_groups`、`user_group_memberships`：部门和管理关系。
- `Presence.form_content`：考勤配置 JSON，包含人员条目、月份、日历例外与合同快照。
- `personnel_contracts`：合同历史。月中合同切换可以形成两份分别绑定有效区间的考勤表。
- `transport_receipts`：文件所属人员、上传者、类型、绑定月份和归档状态；文件内容在私有目录中。
- `leave_requests`、`leave_events`：新版带薪假申请与操作历史；新模块尚未上线，不连接考勤或邮件。注销账户保留此历史，但解绑本人访问关系。

本地重构启动脚本为 `scripts/start_local_rebuild.ps1`，首次快照导入工具为 `scripts/setup_local_snapshot.py`；配置、数据库、来源快照和文件均放在被忽略的 `.local-test/rebuild/`。详情及验收边界见 `local-rebuild.md`。

日历兼容结构同时包含 `vacances`、`absences`、`arret` 和 `calendar_overrides`。不能因界面不再显示旧输入区，就删除旧字段或历史数据。

## 常见修改的最小阅读范围

| 需求 | 先读 |
| --- | --- |
| 改 ZIP 文件名或容量 | `document_downloads.py` |
| 改考勤必填校验或文件命名 | `fiche_exports.py` |
| 改 Excel 样式或单元格 | `ExcelGen.py`、`Fiche_Exemple.xlsx` |
| 改日历状态与导出结果 | `calendar_view.py`、`contract_periods.py`、`ExcelGen.py` |
| 改合同切换 | `contract_periods.py`、`backend/contract_history.py` |
| 改签署人或签署 API | `backend/signature_flow.py`、`backend/esign.py` |
| 改 PDF 签名框位置 | `backend/signature_placement.py`、`backend/signature_flow.py` |
| 改人员或团队权限 | `access_control.py`、`backend/attendance_permissions.py`、`backend/main.py` 中相关授权函数 |

## 权限与文件生命周期

SuperAdmin 继承 Admin，并独占资源永久删除权限；Admin 保留管理、编辑及归档权限。Responsable 可以在后端允许的团队范围内查看并填考勤，不能修改人员、合同或删除考勤；Employe 使用自己的账户和允许的个人资料。Documents 普通员工仅本人/上传者可归档恢复，非 Admin 的 Responsable 保持只读；权限均由实时后端检查执行。

Documents 的文件默认保存在 `/var/lib/presence-app/transport-receipts`，由 `PRESENCE_TRANSPORT_RECEIPTS_DIR` 配置。归档只改变元数据，不删除文件。下载需要认证，Admin 可见全部，普通账户与经理按后端访问范围查看。

供 ClawShow 获取的临时 PDF 默认保存在 `/var/lib/presence-app/esign-documents`，有效期为 24 小时；它与长期保存的员工凭证是不同的存储机制。

## 历史限制清单（含旧版 Streamlit，不能直接作为新版现状）

- ClawShow 已支持通过 `party: 1` 签名字段开启负责人手写。员工沙盒样本已完成双签并验证最终 PDF；用户随后通过本地页面连接正式后端确认真实双签流程跑通。2026 年 10 月 14 日前后端模块化版本已整体部署，公网健康检查通过。实习生完整双签仍待验证。
- Documents 上传目标已限制为本人，适用于员工与经理；Admin 可代传。非 Admin 的所属人员接口只返回本人，后端拒绝将上传目标改为他人。经理继续查看授权团队范围，归档、月份修正及删除权限未扩展；已有文件归属不变。
- Documents 尚无归档恢复操作、一次上传多个文件；考勤尚无文档概览以前描述的 60 天回收站流程。
- 日历部分 `calendar_overrides` 状态与 Excel 导出尚需统一；实习生报酬仍使用手动天数和每日时数。
- 上传控件显示的 200 MB 与后端实际 10 MB 限制尚需对齐。
- 考勤页面的 `fiche_name` 同名变量覆盖已修复，批量下载和签署共用勾选列表；配置读取失败后仍进入编辑流程的问题尚未修复。
- 签署的第二人仍为员工档案中的负责人。当前未接收签署完成 webhook，也没有将双方已签 PDF 自动回存到 Documents 或考勤归档的功能。
- 退出登录未清理文件和考勤缓存；单份生成结果未按内容变化失效。需补齐账号隔离和缓存生命周期。
- 删除账号是永久删除，会清理考勤、合同和相关凭证；包括由该账号上传或归档的他人凭证。当前行为不提供历史数据保留或恢复承诺。
- `backend/main.py` 仍集中管理认证、组织和提醒。后续可按这些边界继续拆分，避免一次改变全部模块。

## 维护顺序

先定位模块，读取该模块与直接依赖；修改共用逻辑时检查所有调用方。记录接口、字段或部署文件变化。静态检查、功能测试及真实上传或邮件发送应分别记录结果。

## 2026-10-15 本地重构补齐

本轮补齐跨模块草稿保护、模块错误隔离与手动重试、新建账户表单、提醒发送/重发确认，以及后端能力开关。154 项 Python 回归与 7 项前端规则测试通过。普通工作流未实际新建账户或发送邮件，保持独立本地测试。发布前配置、尚待验收项与回退边界见 [新版发布准备](rebuild-release-readiness.md)。前述已知限制中针对 Streamlit 的条目，不代表新版 React 同样缺失。

2026-10-15 UI 更新：考勤标题中的人员名称使用适配双主题的背景高亮，静态前端版本 `name-highlight-20261015T103005Z`；后端与数据不变。回退记录见 `rebuild-release-readiness.md`。

共用删除入口：`frontend/src/components/DeleteAction.tsx` 提供 44px 垃圾桶按钮、悬停/聚焦提示和可访问名称，用于考勤、Documents 与 Admin 账户；原权限与确认步骤保留。


2026-10-16 Documents 视觉版（已发布 documents-visual-20261016T095254Z）：降低标题/控件字重，月份工具栏对齐说明；搜索改为带无障碍名称的占位输入；文件夹和类型改为胶囊筛选；标签统一使用主题色半透明底；卡片至少 16px 内边距，预览/下载与更多菜单简化操作。保留当前月份搜索、三档密度、权限及删除确认；未改全局导航和其他页面主题。


个人月度档案第一版：考勤仅计明确 sheet_month/sheet_year 的关联记录，不推断完成或审批；旧版无月份记录单独提示。签署接受回执新增 profile_subject（实际收件邮箱唯一关联的账号、年月），保存在原签署回执 JSON 中，无新增数据库表；只读查询不创建 SQLite。旧回执无法可靠定位账号/月份的不回填，不能据无回执判断未发送；不表示双方签署完成。请假按日期与月份交集展示（含半天和审批状态，不计算余额）。无真实上传、邮件或签署验证，未部署；构建、lint、Python AST 与本地 Admin 浏览器查看通过，其他角色尚未进行完整回归。


2026-10-21 账号保留策略（本地完成，未上线）：Administration 的永久删除入口改为停用/重新启用与最近 100 条状态历史。所有原有账号迁移默认 active/session_version=0，兼容未携带版本的旧会话；状态改变递增版本并撤销记住登录和重置令牌，后续请求按数据库实时状态及版本校验。保留原工资/考勤配置、合同、凭证、部门归属与负责人关系，Admin 可继续查历史。停用负责人不自动转移下属，界面提示人工检查。老删除接口禁用；文件单项删除仍遵循现有规则。

本地验证：生产构建、lint、Python AST、数据库启动增量迁移和 Admin 页面/停用确认窗口只读查看通过；没有执行实际停用或重新启用，没有运行自动化或多角色回归，没有上传/邮件/签署。部署须同时更新所有仍可访问的 API 的认证与旧删除接口，不能只更新前端；详见 backend/README.md。


2026-10-21 账号两步处理（本地，替代前述“仅停用”方案）：先停用，再通过 `DELETE /users/{username}/login` 删除登录能力。后端事务锁后重新确认 Admin、会话版本、非本人、已停用和完整用户名确认；正常账号不能直接删除。增量增加 `users.login_deleted_at` 与状态事件 action；删除时替换密码哈希、撤销所有登录凭证，记录 login_deleted，不物理删除身份行或任何业务记录。已删除登录不可重新启用，用户名/邮箱保留以防档案误关联；Admin 列表保留“Connexion supprimée”历史身份供查询。旧 /delete-user 路径继续返回 409。`AccountAccess.tsx` 仅对停用且尚未删除的账号显示删除图标和用户名二次确认，失败信息在弹窗内显示。未实际停用/删除现有账号，未运行功能回归，未上线。


2026-10-21 账号流程补充验证（替代上文“未运行回归”的状态）：新增 `test_account_lifecycle.py`（24 项，临时 SQLite/隔离 HTTP，检查认证、保留数据、回滚和迁移）与 `test_account_lifecycle_mysql.py`（1 项端到端 MySQL 临时表流程，默认跳过，设置 PRESENCE_MYSQL_TEMP_TEST=1 显式运行）。使用本地 MySQL 3307 的单连接 TEMPORARY 表，关闭连接自动消失，不写入真实 users 或业务表；测试写入保护仅允许临时 users/account_status_events。开启该项后后端 discover 179 项全部通过；前端草稿测试 7 项、build/lint 通过。SQLite 删除/重新启用竞态模拟另外重复 25 次通过，不等于真实 MySQL 多连接负载测试。首轮回归发现旧测试未适配编号 SQL 字段、签署账号解析与 Profile 入口注册；已更新相应模拟/断言，未降低业务断言或修改产品代码以跳过检查。无真实停用、删除、上传、邮件或签署，未部署。
