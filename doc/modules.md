# Presence 模块索引与维护说明

2026-10-15：React 重制版已发布至 `https://app.presence.217.65.145.103.nip.io/`，版本 `react-20261015T101334Z`。新版服务 `presence-react-api` 使用 127.0.0.1:8015；原 Streamlit 与旧 API 保留运行以供回退。服务器使用原线上数据库与文件，未导入本地快照。发布与回退记录见 `rebuild-release-readiness.md`。

本索引用于按功能定位代码，减少每次开发重复阅读整个项目。前端入口是 `online_test.py`，后端入口是 `backend/main.py`。路径均相对于仓库根目录。

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
新版视觉公共组件：`components/DepartmentBadge.tsx` 提供固定列宽、稳定配色、长名称省略与全名提示；`attendance/AccountOption.tsx` 提供部门/姓名分列选人；同一标签共用于团队名单与 Admin 部门选择。`theme.ts` 统一控件高度、表格及折叠区；`index.css` 管理页面间距、人员控件与窄屏月份对齐。
新版配色在 2026-10-15 按用户反馈改为蓝色系：深色模式主色 `#7aa7f8`、浅色模式 `#285fbd`；保留深色背景 `#0e1117` 与灰色控件 `#262730`。姓名高亮、常规按钮、标签页与日历选择框使用主题主色；错误和删除继续采用语义红色。
新版日历拖选由 `frontend/src/attendance/useCalendarSelection.ts` 管理：鼠标/触摸 Pointer Events 沿经过的半天格添加或取消选择，支持键盘点击；`AttendanceCalendar.tsx` 仍负责显示和统一应用状态。
新版日历状态颜色集中在 `frontend/src/attendance/calendarStatus.ts`，共用于半天格、状态选择器及图例，分别适配 Light/Dark；拖选使用主题色边框保留状态底色。成功提示统一由 `theme.ts` 使用中性灰色呈现。

## 页面与共用模块

| 功能 | 入口 | 共用模块与职责 |
| --- | --- | --- |
| 导航与登录状态 | `online_test.py` | `access_control.py` 解析前端角色；`remember_me.py` 同步持久登录令牌 |
| 登录及找回密码 | `pages/Login.py` | `api_client.py` 提供 API 地址、认证头、请求异常和错误信息解析 |
| 账户设置 | `pages/Profile.py` | 调用后端资料更新接口 |
| 考勤表创建与编辑 | `pages/Fiches.py` | 组织人员选择、合同绑定、分页、保存与导出操作 |
| 考勤校验与生成 | `fiche_exports.py` | `fiche_name`、`validate_fiche`、`generate_fiche_file`，供单份、批量导出和批量签署共用 |
| 考勤批量选择 | `fiche_batch_ui.py` | 按账号和月份隔离勾选状态，提供全选、取消全选；下载与签署仅处理勾选的表单 |
| 考勤日期序列化 | `fiche_serialization.py` | 保存时日期转 ISO 字符串；读取时恢复日期对象 |
| 默认工作安排 | `planning.py` | 工作日名称、默认时段及新记录的默认安排；不覆盖已保存安排 |
| 日历显示与修改 | `calendar_view.py` | 半天状态解析、例外记录同步、批量编辑与预览 |
| 日历浏览器组件 | `calendar_selector_component/index.html` | 日历交互与拖动选择；组件加载由 `calendar_view.py` 管理 |
| 月份选择 | `period_selector.py` | 法语月份名和同一页面月份选择器之间的状态联动 |
| 合同与月份匹配 | `contract_periods.py` | 合同快照、月份内有效区间、跨月例外过滤及合同文件名 |
| 合同编辑界面 | `contract_ui.py` | 管理员维护合同的共用界面 |
| Excel 模板生成 | `ExcelGen.py` | 将工资与考勤数据写入 `Fiche_Exemple.xlsx` |
| 实习生 Word 生成 | `DocxGen.py` | 使用 `template_stagiaire.docx` 生成文件 |
| 前端 PDF 转换请求 | `pdf_export.py` | 调用 `/fiches/pdf` 并检查返回内容 |
| Documents 页面 | `pages/TransportReceipts.py` | 上传表单、顶部与文件列表上方联动的月份选择、横向标题和过滤栏、待处理或归档视图、三列卡片 |
| Documents 卡片与操作 | `document_ui.py` | 月份显示、下载准备、归档、删除确认、月份修正、勾选和 ZIP 下载控件 |
| Documents ZIP 内容生成 | `document_downloads.py` | 逐个鉴权下载、文件名清理、100 MB 上限；失败时不返回部分 ZIP |
| 人员及组织管理 | `pages/Admin.py` | 账户、角色、部门、负责人及合同管理 |
| 团队查看 | `pages/People.py` | Mon équipe：架构默认收起；员工查看本组基础名单，经理与 Admin 的原表格保留 |
| 团队数据请求与名单 | `team_api.py`、`team_ui.py` | 带认证读取与表格整理；普通员工本组名单与管理人员原授权名单分开 |
| 部门架构 | `organization_model.py`、`organization_ui.py` | 纯数据投影与折叠展示分开，无示例数据和预览控制；`organization_chart.py` 仅保留旧导入兼容 |
| 响应式部门卡片 | `organization_cards.py` | 部门独立卡片与 CSS Grid 自动换行，Direction Générale 位于顶部；部门之间无连线，无拖动、缩放或固定画布高度 |
| 提醒及仪表板 | `pages/Dashboard.py` | 合同到期提醒、年度面谈与设置展示 |
| 新版带薪假申请 | `frontend/src/App.tsx` | React + TypeScript + Vite + Material UI；个人/团队申请、分页、过滤、系统内提醒、申请/审批窗口及操作历史，接口集中在 `src/api.ts` |
| 新版总入口及请假页面 | `frontend/src/AppWorkspace.tsx`、`frontend/src/LeaveWorkspace.tsx` | `App.tsx` 只管理会话，外壳提供统一导航，旧请假页面独立嵌入；界面主题与角色菜单共用 |
| 新版页面加载和草稿提醒 | `frontend/src/AppWorkspace.tsx` | 各模块独立加载与失败重试，已打开模块保留状态；`draftGuard.ts`、`draftRules.ts`、`components/DraftGuardProvider.tsx` 共用草稿注册、确认与刷新/关闭提醒，覆盖考勤、管理表单、Documents 上传、请假、面谈、账户设置；提交期间阻止工作区切页/退出 |
| 新版考勤页面 | `frontend/src/attendance/AttendancePage.tsx` | 保留完整 JSON 草稿、稳定 ID 分页、保存、导出与勾选；保存固定在当前考勤页面右下角，单份 Excel/Word 和 PDF 收进 `FicheExportMenu.tsx` 的 More 菜单；编辑、日历及月份拆至同目录组件，API 复用总请求工具 |
| 新版 Documents | `frontend/src/documents/DocumentsPage.tsx` | 月份联动、待处理/归档、类型筛选、勾选及批量下载；`DocumentUpload.tsx` 上传，`DocumentCard.tsx` 卡片及操作确认，`api.ts` 类型与请求；共用 `src/download.ts` 下载工具 |
| 新版 Mon équipe | `frontend/src/team/TeamPage.tsx` | 以横向标签页切换人员名单和公司架构，默认显示名单；按角色请求本组目录或授权员工列表，支持搜索且切页保留搜索内容；`OrganizationCards.tsx` 在架构标签页直接渲染响应式部门卡片，`api.ts` 定义只读请求与类型；服务端复用 `organization_model.py` 投影 |
| 新版 Administration | `frontend/src/admin/AdminPage.tsx` | Admin 入口、账户搜索/选择、操作状态和查询更新；`CreateAccount.tsx` 新建账户表单与 `UserEditor.tsx` 共用 `OrganizationFields.tsx`；`UserEditor.tsx` 角色/部门/负责人及关联人员，`ContractEditor.tsx` 合同历史，`GroupEditor.tsx` 部门创建/修改，`api.ts` 请求类型；复用原管理接口，创建功能由 `/ui-capabilities` 控制，本地后端仍禁止新增账户 |
| 新版能力配置 | `frontend/src/capabilities.ts`、`backend/ui_capabilities.py` | `GET /ui-capabilities` 只返回环境与布尔功能开关；界面无法核对时禁用新增账户和邮件，权限仍由各业务接口执行；本地入口另外拦截创建、找回邮件与提醒发送/重发 |
| 新版仪表板 | `frontend/src/dashboard/DashboardPage.tsx` | 合同到期、面谈统计及名单、Admin 提醒偏好及发送/重发确认窗口；`InterviewEditor.tsx` 编辑面谈，`api.ts` 共用原接口；本地禁用真实提醒发送 |
| 新版账户设置 | `frontend/src/profile/ProfilePage.tsx` | 本人用户名/邮箱/密码修改及当前密码确认；成功后交给 `App.tsx` 更新会话令牌、清理查询缓存并重载账户状态 |
| 新版持久登录与密码重置 | `frontend/src/authSession.ts`、`components/PasswordRecovery.tsx` | 可选记住登录、恢复/撤销；访问令牌保存在 sessionStorage、持久随机令牌保存在 localStorage，密码不持久化；支持 `?token=` 重置入口，本地不发送找回邮件 |
| 新版考勤签署 | `frontend/src/attendance/SignaturePreview.tsx`、`backend/attendance_signatures.py` | 单份/勾选人员预检、确认及逐份发送；直接复用旧版签署服务。`signature_attempts.py` 用本地 SQLite 记录提交回执，防同一次请求重复发送；仅沙盒配置就绪时启用 |
| 签署链接与耗时 | `backend/esign.py`、`frontend/src/attendance/SignaturePreview.tsx` | 沙盒只展示返回的签署页，不展示 PDF/区域设置页；依响应中的邮箱标注签署人，无法核实时保留通用标签；记录生成/PDF/服务商耗时，前端显示等待秒数与明确完成状态 |
| 本地 PDF 专用公网入口 | `backend/pdf_test_gateway.py`、`scripts/start_pdf_test_tunnel.ps1`、`scripts/stop_pdf_test_tunnel.ps1` | 独立 8014 端口，只读取随机令牌对应且未过期的 PDF；Cloudflare 临时隧道只连接此入口，未挂载 8013 私有 API；启动/停止和临时地址写入本地忽略目录 |
| 指定人员正式邮件测试 | `backend/local_signing_policy.py`、`scripts/stop_live_signing_test.ps1` | 默认沙盒；用户明确授权后，可在本地私有配置限时开放唯一员工/负责人邮箱组合；新适配器与旧发送服务均校验，到期拒绝发送，支持撤销并恢复沙盒 |

## 后端模块

| 模块 | 职责 | 主要接口 |
| --- | --- | --- |
| `backend/main.py` | 应用入口、环境配置、数据库连接、认证、组织权限、用户、提醒、面谈、考勤 JSON 保存和启动迁移 | `/login`、`/restore-session`、`/groups`、`/list-users`、`/dashboard`、`/get-config/{user_id}`、`/save-config/{user_id}` 等 |
| `backend/documents.py` | 文件上传、访问范围、元数据、月份修正、归档及删除；ZIP 逐份复用下载鉴权，最多 100 份、100 MiB，不返回部分 ZIP | `/transport-receipts`、`/transport-receipts/assignees`、`/transport-receipts/{receipt_id}/file`、`/period`、`/archive`、`POST /transport-receipts/export` |
| `backend/esign.py` | LibreOffice 转换、临时 PDF 存储、ClawShow 请求及接口错误转换；注册函数返回原发送服务供 React 适配器复用，原路由不变 | `/fiches/pdf`、`/send-fiche`、`/esign/documents/{document_token}` |
| `backend/signature_flow.py` | 根据员工账户寻找负责人，检查签署邮箱与权限，组装 SIGNER 顺序及员工或实习生的两个签名锚点 | 由 `backend/esign.py` 调用 |
| `backend/signature_placement.py` | 从最终 PDF 查找唯一签名文字，按 ClawShow 当前的 A4 参考坐标换算；拒绝缺失、重复及不支持的旋转或裁剪页面 | 由 `backend/esign.py` 调用 |
| `backend/contract_history.py` | 合同历史表创建与旧数据迁移，维护合同并阻止日期重叠 | `/users/{username}/contracts`、`/users/{username}/contracts/{contract_id}` |
| `backend/attendance_permissions.py` | 限制经理保存内容为权限范围内的考勤字段 | 由 `/save-config/{user_id}` 调用 |
| `backend/team_directory.py` | 登录用户查看基础架构；目录响应新增 `organization` 投影供 React 使用，保留旧字段；按服务端部门归属查询员工本组名单，不暴露合同、完整邮箱或财务资料 | `/organization-directory`、`/my-team` |
| `backend/leave_requests.py` | 带薪假接口，注入认证/连接；跟随当前直接负责人，Admin 代审，员工审批前撤回 | `/leave/session`、`/leave-requests`、`/{id}`、`/{id}/decision`、`/{id}/withdraw` |
| `backend/leave_rules.py`、`backend/leave_store.py` | 日期/半天/权限与存储分离；事务内保存状态和操作历史，建表；账号关联字段排序规则按 `users.username` 对齐，保留旧数据；入口账号改名/删除时同步或解绑申请归属 | 正式 MariaDB；独立本地预览使用 SQLite |
| `backend/leave_preview.py` | 本地虚构身份选择、独立预览库，不能部署到生产 | `/preview/accounts`、`/preview/login`；共用正式带薪假路由 |
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
| 调整文件卡片、勾选或显示 | `document_ui.py` |
| 改上传类型或绑定月份 | `pages/TransportReceipts.py`、`backend/documents.py` |
| 改 ZIP 文件名或容量 | `document_downloads.py` |
| 改考勤必填校验或文件命名 | `fiche_exports.py` |
| 改 Excel 样式或单元格 | `ExcelGen.py`、`Fiche_Exemple.xlsx` |
| 改日历状态与导出结果 | `calendar_view.py`、`contract_periods.py`、`ExcelGen.py` |
| 改合同切换 | `contract_periods.py`、`backend/contract_history.py` |
| 改签署人或签署 API | `backend/signature_flow.py`、`backend/esign.py` |
| 改 PDF 签名框位置 | `backend/signature_placement.py`、`backend/signature_flow.py` |
| 改人员或团队权限 | `access_control.py`、`backend/attendance_permissions.py`、`backend/main.py` 中相关授权函数 |

## 权限与文件生命周期

Admin 拥有管理权限；Responsable 可以在后端允许的团队范围内查看并填考勤，不能修改人员、合同或删除考勤；Employe 使用自己的账户和允许的个人资料。

Documents 的文件默认保存在 `/var/lib/presence-app/transport-receipts`，由 `PRESENCE_TRANSPORT_RECEIPTS_DIR` 配置。归档只改变元数据，不删除文件。下载需要认证，Admin 可见全部，普通账户与经理按后端访问范围查看。

供 ClawShow 获取的临时 PDF 默认保存在 `/var/lib/presence-app/esign-documents`，有效期为 24 小时；它与长期保存的员工凭证是不同的存储机制。

## 已知限制

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
