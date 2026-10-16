# React 新版发布记录与剩余验收

2026-10-15：React 重制版已上线。本地测试环境继续独立运行，未将测试数据导入线上。

## 本次发布

- 版本：`react-20261015T101334Z`。
- 网址：`https://app.presence.217.65.145.103.nip.io/`。
- 新服务：`presence-react-api`，127.0.0.1:8015，已设开机启动。
- 代码：`/var/www/presence-react/releases/react-20261015T101334Z`，`/var/www/presence-react/current` 指向该目录。使用完整仓库入口 `backend.rebuild_server:app`。
- 备份：`/var/backups/presence-react/react-20261015T101334Z`，包含数据库转储、原服务代码、Nginx/systemd/环境配置和私有文件；仅 root 可读取。
- 旧版：`presence-interface`（8501）和 `presence-app`（8001）均保留；没有覆盖其代码或环境文件。
- 自动提醒：由旧 API 保留唯一后台任务；新服务设置 `PRESENCE_CONTRACT_REMINDER_WORKER=0`，手动提醒功能仍按权限可用。以后停用旧 API 时需把新服务该开关改为 1，再重启新服务，避免自动提醒停发。
- 数据：线上原库和私有目录继续使用，账户/考勤/合同/部门/成员关系/凭证元数据在部署前后指纹一致。只新增空的请假申请和事件表。

## 回退操作

在服务器执行：

```sh
sudo /var/backups/presence-react/react-20261015T101334Z/rollback.sh
```

脚本核对当前发布标记，恢复原 Nginx 页面入口，确保旧服务运行，再停用新服务。脚本语法与旧版健康已验证，未为演练而实际回切。**不会恢复数据库快照或删除已上传文件**，避免丢失上线后的操作。新版请假历史保留在数据库，旧界面不提供该模块；回退期间应暂停人员改名/删除，避免旧接口无法同步新请假关联。

## 本次验证

- 154 项 Python 测试、7 项前端测试、生产构建和 Oxlint 通过；增加提醒 worker 配置后，服务器入口与 readiness 的 5 项针对回归也通过。
- 切换前后各完成 7 个现有账户的 70 次只读请求检查，服务端原鉴权继续生效。
- 公网页面、JS/CSS 和能力接口 200；未认证的考勤/文件请求 401；本地环境接口在线上 404。
- 4 份现有私有文件成功通过认证下载。服务器 LibreOffice 转换、临时诊断 PDF 的公网 HTTPS 路径与字节一致性验证通过，诊断文件已删除。
- 正式签署预检可用，已有资料不足的历史 fiche 仍被校验拦截。本轮没有发起真实签署、发送测试邮件、重置密码、创建账户或提交请假。
- 公网浏览器登录页正常，未记录启动错误；真实邮件投递与实习生双签仍由实际业务验收。

## 已补齐的迁移入口

考勤编辑/文件生成/批量下载及双签、Documents、Mon équipe、Administration、仪表板、账户设置和找回密码界面已存在。创建账户表单复用旧 `/create-user`；提醒发送/重发复用旧接口。界面显隐之外，仍执行原接口鉴权。

`GET /ui-capabilities` 是公开的布尔能力查询，不返回配置值：

| 字段 | 本地副本 | 服务器配置 |
| --- | --- | --- |
| `environment` | `local-snapshot` | `server` |
| `account_creation` | false | true，提交仍需 Admin |
| `mail_enabled` | false | SMTP 必填配置存在时 true；不代表邮件投递已验证 |

读取失败时，前端禁用上述功能。即使编译为正式静态文件，只要连接本地后端，创建和邮件仍被禁用。本地中间件拦截创建、找回邮件与提醒发送/重发；eSign 使用自己的独立模式和发送策略。

## 发布配置与操作顺序（本次已执行）

1. 对比本地与服务器代码差异，按 `backend/README.md` 备份源码、数据库和私有文件，保留原服务作为回退。
2. 为新版配置独立静态目录与同源 `/api/` 反向代理；不要将开发服务器或本地数据库暴露到公网。
3. 已准备独立入口 `backend.rebuild_server:app`，从完整仓库根目录运行，注册新版考勤及签署适配器，复用原鉴权、版本检查、发送服务；使用 `PRESENCE_SIGNATURE_ATTEMPTS_DB` 保存签署提交回执，默认 `/var/lib/presence-app/signature-attempts.sqlite3`，服务账号需可写。不能把 `local_rebuild.py` 搬到服务器。
4. 服务器 readiness 已支持 `sk_test_` 沙盒与 `sk_live_` Production，校验 HTTPS 地址并在前端显示正式邮件/额度确认。不得复用临时 Cloudflare 测试地址或本地限时测试许可。配置样例位于 `deploy/`，本次按实际目录、账号和端口安装。
5. 配置服务器应用地址供找回密码邮件指向新版 `?token=` 页面；按原配置保留数据库、SMTP、LibreOffice、公开 PDF 地址及私有凭证目录。密钥只放服务器环境文件，不写前端。
6. 环境标签已按后端区分本地副本与服务器，其余操作消息使用中性措辞。发布时核对会话续期、反向代理、HTTPS、静态资源缓存策略。旧资源应保留到现有会话结束，避免更新后分块文件丢失。
7. 用户确认本地验收通过并明确要求发布后，再执行备份、部署、健康检查和角色回归。不要覆盖线上数据为本地快照。

## 仍保留的验收边界

- 自动检查：154 项 Python 测试、7 项前端规则测试、类型检查、构建与 lint。运行 `npm --prefix frontend run test`；Python 使用项目 `.venv`。
- 浏览器检查：请假窗口取消保护、跨模块部门草稿保留、本地新建按钮禁用、最终打包登录页面。
- 仍需用户验收：手机/平板实际拖选与上传、实习生完整双签，以及生产配置下的建号、密码重置邮件与合同提醒投递。此次没有代用户发送这些邮件或创建真实账户。
- 仍有主包体积警告。保留页面级加载与模块手动重试；模块重建可能丢失该模块自身的草稿，其他模块不随之刷新。未承诺浏览器崩溃后恢复草稿。
- 按用户 2026-10-15 决定，暂不开发最终已签 PDF 回存与自动归档，通过 ClawShow 邮件链接下载；此项不作为当前重构待完成项。

部署样例 `deploy/presence-react-api.service.example` 使用独立 loopback 8002 端口及完整仓库；`deploy/nginx-react.locations.example` 提供静态文件、同源 `/api/`、上传大小及签署超时配置。路径、运行账号、HTTPS 主机和公开 PDF 前缀需由实际服务器确定；不要原样替换当前 Nginx。新入口仍会沿用原主服务启动迁移与提醒任务，首次启动也必须按发布操作对待。

## 姓名高亮 UI 更新（2026-10-15）

版本 `name-highlight-20261015T103005Z` 已同步线上，仅更新静态前端，不重启后端、不写数据库。考勤标题中 Fiche de 后的人名使用红色系背景、边框及内边距，适配 Light/Dark 和长名称换行。当前静态根目录为 `/var/www/presence-react/frontend-releases/name-highlight-20261015T103005Z`；保留前版资源以兼容已打开的页面。

仅回退本次 UI：`sudo /var/backups/presence-react/name-highlight-20261015T103005Z/rollback-ui.sh`。原完整回退脚本仍然有效。生产构建、lint 通过；公网首页及考勤 JS 与本地构建字节一致。

## 蓝色主色更新（2026-10-15）

版本 `calm-blue-20261015T103218Z` 已发布：普通强调、操作按钮、选中标签、日历选择边框及姓名高亮统一使用双主题蓝色，避免与错误提示混淆；删除/错误保留红色，各部门和考勤状态分类颜色不变。仅替换静态资源，构建与 lint 通过，公网首页和考勤 JS 字节核对通过。

本次 UI 回退：`sudo /var/backups/presence-react/calm-blue-20261015T103218Z/rollback-ui.sh`。前版资源和完整旧版回退仍保留。

## 删除入口简化（2026-10-15）

静态版本 `delete-icon-20261015T103931Z` 已上线。考勤、文件、账户的删除入口统一改为垃圾桶图标，悬停/聚焦显示相应法语提示，保留原有权限和删除确认；确认窗口最终操作仍显示文字。构建与 lint 通过，公网首页和共用图标组件字节核对通过；未执行任何删除。回退：`sudo /var/backups/presence-react/delete-icon-20261015T103931Z/rollback-ui.sh`。

### 2026-10-15 考勤页签与法语日期

- 更新前 Git 保存点：`8eb4f2c`，标签 `checkpoint/before-attendance-sheets-20261015`。这是当时本地代码快照，密钥、私有运行配置和编译缓存未纳入。
- 前端静态版本 `attendance-sheets-20261015T105159Z`：三个编辑区域改为独立页签，默认日历；法语日期控件替代受系统语言影响的原生日期控件，API 保留 ISO 值。
- TypeScript/Vite 构建及 oxlint 通过，本地只读查看页签与法语日期；线上 HTML、考勤和日期模块与本地构建逐字节一致。未执行业务写入、邮件或签署测试。
- 仅发布前端，后端未重启。上次静态版本保留；本次界面回退：`sudo /var/backups/presence-react/attendance-sheets-20261015T105159Z/rollback-ui.sh`。

### 2026-10-15 Documents 月份栏简化

- 静态版本 `documents-period-20261015T123341Z`：删除文件列表中的重复月份选择，仅保留顶部单一吸顶月份栏，同时控制上传与查询；上传绑定月份提示保留。
- 月份修改继续经过草稿确认，清空旧月份勾选；筛选、归档和权限逻辑不变。
- 构建与 oxlint 通过；线上首页、Documents 模块与本地构建逐字节一致。未上传文件、发送邮件或修改业务数据，后端未重启。
- 上次版本保留；界面回退命令：`sudo /var/backups/presence-react/documents-period-20261015T123341Z/rollback-ui.sh`。

### 2026-10-15 创建账户弹窗

- 前端静态版本 `create-account-dialog-20261015T125621Z`：Administration 顶部按钮打开创建账户窗口，替代原展开区；复用原创建接口与字段。
- 取消、Esc 与点击遮罩关闭检查窗口草稿；提交期间阻止关闭，失败在窗口内显示错误，成功后关闭。保留环境能力控制及后端鉴权。
- TypeScript/Vite 构建与 oxlint 通过，线上 HTML 与 Admin 模块逐字节核对一致；没有创建账户或发送邮件，后端未重启。
- 回退：`sudo /var/backups/presence-react/create-account-dialog-20261015T125621Z/rollback-ui.sh`，上次静态版本保留。

### 2026-10-15 Documents 上传编号

- 版本 `document-numbers-20261015T131336Z`，新增 `backend/document_numbers.py`，当前 React API 与保留的旧 API 同时接入；部署前结构比对确认旧 API 仅缺少新版 ZIP 路由，补丁保留该差异。
- `YYYYMMDD-0001` 按 Europe/Paris 的实际上传日分配，每日全公司共享序列；四位起步，超过 9999 自动扩位。新增唯一列与每日计数表；分配与上传元数据插入同事务，行锁串行化同一天的并发申请。删除不回收序号，修改报销月份及归档不改变编号。
- 已有仍保留的 4 份文件按上传时间/id 补号：20261008-0001、0002、0003，以及 20261009-0001。已经删除且没有历史记录的文件无法追溯补号。
- 编号展示在卡片中，单份下载和 ZIP 文件名包含编号，不修改原始文件名及私有存储路径。
- 后端文件/数据库/私有文件备份：`/var/backups/presence-react/document-numbers-20261015T131336Z-backend`。原文件字节和原有记录字段前后完全一致；公开认证 Documents API 返回 200，4 份均带编号。Python 语法、前端构建与 oxlint 通过；未实际新增上传或执行并发压力测试，未发送邮件。
- 首次发布因 FastAPI 启动钩子名称不兼容自动回退，修改为 router 注册后重新发布成功；两个 API 服务均 active。
- 前端回退：`sudo /var/backups/presence-react/document-numbers-20261015T131336Z/rollback-ui.sh`；后端回退：`sudo python3 /var/backups/presence-react/document-numbers-20261015T131336Z-backend/rollback-backend.py`。回退保留增量编号列及计数表，不覆盖后续上传数据；旧版本上传的空编号在重新启用编号版本启动时补齐。

### 2026-10-15 Documents 在线预览

- 静态版本 `document-preview-20261015T133312Z`：文件卡片增加 Aperçu，弹窗显示 PDF/JPG/PNG，保留编号、下载、新标签页和关闭操作，小屏全屏。
- 复用原带令牌下载接口，响应以 Blob 留在内存；PDF.js 延迟加载、独立 worker、Canvas 渲染，支持逐页与 50–200% 缩放。图片用内存 data URL，关闭时取消未完成请求、销毁 PDF worker 并回收 URL。
- 保留服务器 CSP（object-src none、img-src self/data），没有改为公网文件地址或外部预览服务。Vite 把 PDF.js 6.4.299 的 CMap、字体、WASM/ICC 资源打包到版本目录，生成文件忽略入 Git；构建需要 Node >=22.13（或 >=24）。
- 构建、oxlint 通过；本地已打开现有 Test-PDF.pdf，正文实际渲染；未新增上传，也未改动文件、数据库或权限。线上 HTML、Documents/PdfPreview、worker、CMap、字体及 WASM 抽样与本地字节一致，worker 返回正确 JavaScript MIME。尚未逐一验证所有移动浏览器或复杂 PDF。
- 只发布前端，未重启后端；界面回退：`sudo /var/backups/presence-react/document-preview-20261015T133312Z/rollback-ui.sh`。

### 2026-10-15 Documents 操作图标

- 静态版本 `document-icons-20261015T133934Z`：卡片预览、下载、改月、归档分别用眼睛、下载、铅笔和归档盒图标，沿用垃圾桶删除；统一 44px 点击区域、法语 Tooltip 和 aria-label。
- 原操作及权限保持不变。构建及 oxlint 通过，线上首页和 Documents 模块与本地构建字节一致；无业务写入，后端未重启。
- 回退：`sudo /var/backups/presence-react/document-icons-20261015T133934Z/rollback-ui.sh`。

### 2026-10-15 未保存提醒

- 静态版本 `unsaved-reminder-20261015T134848Z`：复用现有草稿注册与关闭确认，上层统一显示法语待保存模块提示，仅传递状态标签，不记录表单值。考勤右下角显示待保存/待应用日历/无待修改状态，有字段修改才突出保存按钮。
- 日历未应用选区也登记为草稿；切月、切员工或打开其他档案先确认，取消时保留选区。仅选区未应用时禁用普通保存，提示在日历点击 Appliquer；切内部 Sheet 保留选区。已有切模块、退出、刷新/关闭提醒继续有效，保存成功后字段提示消失，待应用选区继续提示。
- 构建及 oxlint 通过；线上 HTML 与考勤模块逐字节核对一致。未执行保存、考勤应用或其他业务写入；仅前端发布，后端未重启。
- 回退：`sudo /var/backups/presence-react/unsaved-reminder-20261015T134848Z/rollback-ui.sh`。

### 2026-10-15 批量操作结果反馈

- 静态版本 `batch-feedback-20261015T135445Z`：考勤导出与 Documents ZIP 复用 `BatchExportResult.tsx`，显示准备中/完整文件已生成并触发下载/中断无部分下载；保留本次月份、选择快照和错误，可下载 JSON 明细。失败保留勾选，不推断某个文件已经成功。
- 签署预检显示选中、有效及待修复数量；处理时显示当前员工、已处理进度和秒数。结果区分已创建（有 accepted 回执及引用）、校验不通过未发送、提交异常待核查；异常可能已受理，不自动重发。结果明细可下载，不导出带权限的签署链接；发送期间接入全局关闭提醒，页面卸载后不继续提交剩余人员。
- 构建和 oxlint 通过，线上 HTML/考勤/Documents 模块与本地构建字节一致；没有实际创建签署请求、发送邮件、导出真实业务批次或执行保存。仅前端发布，后端未重启。
- 回退：`sudo /var/backups/presence-react/batch-feedback-20261015T135445Z/rollback-ui.sh`。

### 2026-10-15 ChristopheH 管理员权限事件

- 用户报告管理菜单消失；只读数据库核对确认 ChristopheH 的 role/role_tags/is_admin 均为普通员工值，非前端缓存。当天存在多次对该账户 organization 接口的成功 PUT 日志；此前未记录操作者，无法确认具体操作人或仅凭访问日志还原每次请求内容。个人资料更新代码不写角色字段。
- 在核验用户名与邮箱后，仅恢复原有 Admin + Employe 权限及兼容 is_admin/role 字段，未改密码、合同、部门或其他账户；变更前备份 `/var/backups/presence-react/admin-incident-20261015T152544Z/role-before.json`。
- 修复同一组织更新端点：事务锁串行化角色编辑并重新读取操作者权限，拒绝自我移除 Admin；提交后的角色改变记录 actor/target/before/after（不记录凭据）。当前 API 与保留的旧 API 只替换该函数，保留服务器其他差异；源文件和数据库备份 `/var/backups/presence-react/admin-guard-20261015T152805Z`。两个服务已恢复 ready。
- 前端版本 `admin-guard-20261015T152823Z`：本人 Admin 项不能取消，其他角色变更在提交前展示新旧角色确认。构建/oxlint 与 Python 语法检查通过；公开认证 GET /leave/session 返回 Admin、Employe，GET /groups 返回 200，未执行降权请求测试。前端资源逐字节核对一致。
- 用户现有会话可能仍显示旧菜单，应刷新或退出重登。前端回退：`sudo /var/backups/presence-react/admin-guard-20261015T152823Z/rollback-ui.sh`；不建议回退后端防误降级修复。

### 2026-10-15 Équipe 人头图标

- 前端版本 `team-person-icon-20261015T153021Z`：部门卡片 Équipe 标签前增加 18px 装饰性人头图标，统一垂直对齐，保留文字。无权限或业务行为改动。
- 构建及 oxlint 通过，线上首页和 Team 模块与本地构建字节一致；后端未重启。
- 回退：`sudo /var/backups/presence-react/team-person-icon-20261015T153021Z/rollback-ui.sh`。


## 2026-10-16 Documents 搜索与卡片大小发布

前端版本 `documents-search-density-20261016T091655Z`：新增当前月份/文件夹内的文件名、编号、员工、上传人及类型说明搜索；新增 Compact/Standard/Confort 卡片密度与浏览器偏好保存。同时包含已提交的前端旧预览清理。后端清理仅在 Git 中，本次不覆盖服务器 Python、不重启后端、不修改数据库。

发布前构建、lint 与差异检查通过。线上原静态根为 `team-person-icon-20261015T153021Z`，保留其全部资源，并保留旧哈希资源以兼容已打开页面。Nginx 配置检查成功后 reload。公网 HTML、入口资源、Documents 模块与 PDF worker 共 7 个资源和本地构建逐字节一致；新旧后端、旧 Streamlit 及 Nginx 均 active。未执行真实上传、邮件或签署，也未做登录后的浏览器交互回归。

回退：`sudo /var/backups/presence-react/documents-search-density-20261016T091655Z/rollback-ui.sh`。脚本仅在当前静态根仍为本版本时恢复旧 Nginx 配置，不操作数据库。浏览器刷新后可使用新功能；未保存表单请先保存。
