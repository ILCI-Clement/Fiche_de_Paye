# ILCI · Espace RH 后端

## 运行入口

| 环境 | 入口 / 端口 |
| --- | --- |
| 生产 React | 仓库根目录运行 `backend.rebuild_server:app`，8015 |
| 保留 API | `/var/www/presence-app` 内运行 `main:app`，8001 |
| 独立本地副本 | `backend.local_rebuild:app`，8013；MySQL 3307 |

生产地址：https://app.presence.217.65.145.103.nip.io/ 。

React 入口复用 `main.app`，注册考勤、签署、月度 Profile、员工名册与资源维护。路由模块通过注册函数接收依赖，不反向导入主入口。flat 共享模块和根目录生成器须一并发布，部署后用实际服务账号验证读取权限与导入。

本地启动使用 `scripts/start_local_rebuild.ps1`，真实数据副本及凭据位于被忽略的 `.local-test/rebuild/`。本地禁止创建虚构账户和普通通知邮件；限时正式签署测试由独立策略控制。不要把本地入口、配置、数据库或密钥部署到服务器。

代码定位见 [模块索引](../doc/modules.md)。

员工名册 Excel 导出已上线：运行时须同时包含 `personnel_export.py` 与仓库
`template/Exemple-de-modele-de-registre-unique-du-personnel.xlsx`。仅有效管理员可
导出，按当前筛选的全部分页、每段任职一行，包含社保号明文并逐人记录审计；
NIR 关闭或密钥无效时拒绝导出。接口与映射见 [名册说明](../doc/personnel-register.md)。

## 当前版本与回退

- 版本：`signed-archive-20261023T111516Z`；Git 提交见服务器 `git-release.json`。
- 应用链接：`/var/www/presence-react/current`。
- 实际目录：`/var/www/presence-react/releases/signed-archive-20261023T111516Z`。
- root 专用备份：`/var/backups/presence-react/signed-archive-20261023T111516Z`。
- 完整验证与已知问题：[发布记录](../doc/releases/signed-archive-20261023.md)。
- 社保号配置：已单独启用生产加密录入/查看，仅管理员可用；[密钥与开关记录](../doc/releases/nir-enable-20261022.md)。应用源码仍为上述版本。

安全回退界面：

```sh
sudo /var/backups/presence-react/signed-archive-20261023T111516Z/rollback-ui.sh
```

脚本恢复上一版静态界面，保留新版 API 的账号状态/会话版本保护及业务数据。不要直接运行不识别 `is_active`、`session_version`、`login_deleted_at` 的旧 API；它可能重新放行停用或删除登录的身份。数据库恢复须单独决定，不能覆盖发布后的业务写入。

旧发布记录见 [历史时间线](../doc/releases/20261014-20261021.md)。当前功能及测试结果见本次发布记录。

## 生产配置

凭据只放 `/etc/presence-app/presence-app.env`，由 systemd 加载；不写入源码、文档、前端或命令输出。

| 变量 | 用途 |
| --- | --- |
| `PRESENCE_API_TOKEN` | 会话签名密钥 |
| `PRESENCE_DB_HOST`、`PRESENCE_DB_USER`、`PRESENCE_DB_PASSWORD`、`PRESENCE_DB_DATABASE` | MariaDB |
| `PRESENCE_SMTP_SERVER`、`PRESENCE_SMTP_PORT`、`PRESENCE_SMTP_USER`、`PRESENCE_SMTP_PASSWORD` | 通知邮件 |
| `PRESENCE_APP_URL` | 账号/重置邮件的前端地址 |
| `PRESENCE_TRANSPORT_RECEIPTS_DIR` | 私有文件目录；默认 `/var/lib/presence-app/transport-receipts` |
| `PRESENCE_ESIGN_DOCUMENTS_DIR` | 临时签署 PDF；默认 `/var/lib/presence-app/esign-documents` |
| `PRESENCE_ESIGN_PUBLIC_URL` | ClawShow 可访问的 HTTPS API 地址，当前为网站 `/api` |
| `PRESENCE_ESIGN_CALLBACK_URL` | 完成回调的完整 HTTPS 地址，生产应为网站 `/api/esign/callback`；与临时 PDF 网关独立配置 |
| `PRESENCE_SIGNED_FICHES_DIR` | 永久私有归档；默认 `/var/lib/presence-app/signed-fiches`，必须与 MariaDB 一起备份 |
| `PRESENCE_SIGNED_FICHE_WORKER` | 归档下载任务默认 1；多进程使用数据库租约，不重复归档；候选只读服务可设 0 |
| `CLAWSHOW_ESIGN_API_KEY`、`CLAWSHOW_ESIGN_NAMESPACE` | 服务商凭据/命名空间，默认命名空间 `ilci` |
| `PRESENCE_ESIGN_ORGANIZATION_NAME`、`PRESENCE_ESIGN_ORGANIZATION_FULL_NAME` | 签署邮件机构名称 |
| `PRESENCE_ESIGN_REPRESENTATIVE_NAME`、`PRESENCE_ESIGN_REPRESENTATIVE_TITLE` | 邮件落款；默认解析负责人和 Responsable |
| `PRESENCE_SIGNATURE_ATTEMPTS_DB` | 持久签署回执 SQLite；当前 `/var/lib/presence-app/react/signature-attempts.sqlite3` |
| `PRESENCE_CONTRACT_REMINDER_WORKER` | React 设为 0，由保留 API 运行唯一自动提醒任务；停用旧服务前须移交任务 |
| `PRESENCE_PERSONNEL_NIR_ENABLED`、`PRESENCE_PERSONNEL_NIR_KEYS_FILE` | 生产 NIR 默认关闭；仅精确开关 1 且独立 keyring 有效时开放 |

依赖见 `requirements.txt`；PDF 转换还需 LibreOffice。AES-256-GCM 使用独立密钥，不从会话密钥派生。POSIX keyring 使用 0600；Windows 须另设 NTFS ACL，密钥独立备份且轮换时保留仍被引用的旧 key。

## 数据与关键行为

密码登录保护（已上线）：同一账户连续 5 次密码错误后暂停密码登录 10 分钟；第 5 次及冷却中的请求返回 HTTP 429 和剩余秒数 `Retry-After`。成功登录清零；冷却到期后的第一次错误从 1 重新计数。未满 5 次时，距上次失败满 10 分钟，原失败序列失效，下一次尝试先按清零处理（再次输错计为第 1 次），无需后台定时任务。期间重复请求不延长冷却，用户名/邮箱及大小写变体共用计数。`users.login_failed_attempts`、UTC `login_locked_until` 与 `login_last_failed_at` 持久保存，按账户行锁串行更新，刷新页面、重启或切换 API worker 不清零；旧版无失败时间的记录从新序列开始，不记录尝试的密码。停用/删除登录仍统一拒绝，已有有效会话与记住登录流程不因此撤销。

React 根据 Retry-After 显示倒计时并暂停提交；后端独立执行限制。部署须同步更新所有仍可访问的 `/login` 后端并执行新增字段迁移，不能仅发布登录界面。已通过独立临时账户的错误计数、别名共用、冷却与闲置重置测试；未锁定真实生产账户。

- 登录接受绑定邮箱或用户名，大小写及标识两端空白不敏感；多身份匹配统一拒绝，密码空白保持原样。
- 账号停用/角色变更递增会话版本并撤销持久登录；删除登录保留身份及业务历史且不可重新启用。物理删除仅 SuperAdmin，先停用、精确确认并处理依赖，不级联删除他人文件。
- 员工名册与登录账号、考勤合同独立。Admin 手动一对一关联；本人只读，直接负责人不因此获得名册隐私。关联版本检查、解除/替换和审计同事务提交，失败回滚。
- `/account-personnel/{username}` 仅本人/Admin；`/choices`、`/link` 仅 Admin。关联提交包含 `employee_id`、`expected_employee_id`、`expected_version`、`target_version`，无关联值为 null。
- 人名按显式关联的名册姓名显示，否则使用有效邮箱前缀；身份和权限仍以账号 ID 判断，历史快照不自动回写。
- 考勤保存按版本和稳定 ID 检查；历史表不因当前合同更改而丢失。新建须有有效合同；Admin 用归档，SuperAdmin 删除还需明确 `deleted_ids`。
- Documents 接受单份 PDF/JPG/PNG，最大 10 MiB。员工/经理只能为本人上传，Admin 可代传；文件私有下载，ZIP 最多 100 份/100 MiB，失败不返回部分结果。
- 文档编号按巴黎上传日分配 `YYYYMMDD-0001`；行锁保证每日序号唯一，历史缺号启动时补齐，归档/改月不改号，删除不回收。
- 签署由员工先签、当前直接负责人后签。LibreOffice 转换后按最终 PDF 定位签名，拒绝不明确位置；临时公开 PDF 默认保留 24 小时。接受回执不代表完成。最终 PDF 回存已上线，见下节。
- 请假按半天申请，审批前可撤回，跟随当前直接负责人或 Admin 审批并留历史；不自动计算余额或写考勤。
- `GET /ui-capabilities` 只返回环境和布尔能力；就绪检查不证明实际邮件投递、外部 PDF 获取或业务权限。

NIR 原号与密文不进入普通 Profile；专用查看须先提交审计，禁止缓存，前端短时显示后清除。更多字段和接口规则见 [员工名册说明](../doc/personnel-register.md)。

## 已签考勤文件自动归档

- 按 [ClawShow 当前文档](https://clawshow.ai/en/api)，创建签署时附上 `callback_url` 和独立的随机 `reference_id`。`POST /esign/callback` 持久接收通知；仅 `document.completed` 或 `signer.signed` 且 `all_signed` 严格为 true 时进入下载队列。归档月份来自发送时的考勤月份，不取完成日期。
- 服务商当前文档未给出 HMAC 头。每次请求生成 256 位随机回调凭据，放在 POST 的 `reference_id` 中，数据库只保存其 SHA-256；同时核对 namespace、provider 和已绑定的 document_id。凭据不放 URL、前端或日志。禁止记录回调请求体。此机制不是服务商数字签名验证；后续若提供 HMAC，应升级接收校验。
- 下载仅允许 `https://esign.clawshow.ai/esign/{document_id}/signed.pdf`，拒绝重定向；限制 25 MiB、下载时长和 PDF 页数。检查 PDF 结构后逐字节保存，不重写已签文件，也不声称已独立验证 PDF 的证书信任链。记录 SHA-256，下载时再次校验完整性。
- 数据在 `signed_fiche_documents`、`signed_fiche_events`；文件路径为私有目录下 `年份/月份/随机档案ID.pdf`。随机文件名避免路径、同名及隐私问题；下载文件名含签署时员工姓名、考勤月份和唯一 ID。状态通知去重且完成后不回退；文件原子创建、不覆盖。
- 后台每 15 秒检查持久队列，行锁及 180 秒租约协调多进程；下载失败从 30 秒起指数退避，最长每小时重试，重启后继续。仅重试文件下载，不重新发送签署邀请。Admin 可手动唤醒下载重试。完成通知先落库再返回 204，服务商失败重投不会重复建档。
- 考勤页新增 `Fiches / Signatures` 切换；文件列表每页 20 项，眼睛预览、下载、状态和恢复下载。个人 Profile 可查看本人或授权人员的月度文件。生产与 SPECIMEN 分开显示和导出。
- `GET /signed-fiches`、`/{id}/file`：仅本人、当前直接负责人（Responsable）及 Admin/SuperAdmin；权限由后端实时核对。`GET /signed-fiches/export.zip` 仅 Admin/SuperAdmin，按月导出所有已签归档 PDF，ZIP 内保持年份/月目录。最多 5 000 份及原始文件合计 100 MiB，待回收或缺失/损坏时整包拒绝，避免不完整结果。下载和导出记审计，响应禁止缓存。
- 改名与归档人员关联在同一数据库事务更新；停用/删除登录不清除档案。永久删除账户前检查已签档案依赖，无级联删除文件入口。
- 未配置 `PRESENCE_ESIGN_CALLBACK_URL` 时保留现有发送能力，但预检及列表明确提示不能自动回收。旧请求没有随机凭据/回调，不能自动补收；不会因此重新发送邀请。

发布时为所有仍可访问的 API 加入共享新模块及两张表迁移，使用同一私有归档路径并保证服务身份有读写权限，确保公网 HTTPS 回调 POST 能到达该 API，不能指向只提供临时 PDF 的网关。本地可在私有 `runtime.json` 的 `esign_callback_url` 配置独立回调地址，绝不复制生产凭据。备份数据库与永久文件；回退 UI 时保留回调和下载 worker，不能回退到会丢弃完成事件的版本。

本地沙盒操作：在仓库根目录运行 `scripts/start_pdf_test_tunnel.ps1`。脚本启动/重载只开放限时 PDF 与单个完成回调的 8014 网关，建立临时 HTTPS 隧道，写入本地的 PDF/回调地址，并重启 8013；不会自动发送签署或改变签署模式。确认预检为 Sandbox 后，通过本地页面创建并依次完成两个签名，在 Signatures 的 Tests SPECIMEN 中查看归档；保持隧道及 API 运行至回收完成。结束后运行 `scripts/stop_pdf_test_tunnel.ps1`。若要先取消正式邮件测试授权，运行 `scripts/stop_live_signing_test.ps1`。网关不代理登录、员工名册、文件列表等私有接口，禁止启用请求体/访问日志。

本版后端 403 项、前端 43 项、lint/类型/生产构建通过；用户已手动完成一份 Sandbox 双签并通过实际回调归档，原始 PDF 大小、SHA-256 和结构核对正确。线上候选 108 项、发布后新旧 API 176 项检查通过，已配置正式回调；没有新增 Production 测试签署或发送正式通知邮件。

## 发布顺序

1. 对比两套线上 API 与本地差异，备份源码、MariaDB、私有文件、签署 SQLite 和服务/Nginx 配置。
2. 将完整依赖放到独立版本目录，保留服务器凭据和文件路径；以实际服务账号检查依赖、flat/package 导入及文件权限。
3. 使用隔离 loopback 候选服务验证启动迁移、角色边界与已有数据指纹。迁移应增量、幂等，不清空历史。
4. 所有仍可访问的 API 同步账号与删除保护，再切换前端；保留上一版静态文件和旧哈希资源。
5. 检查新旧 API、Streamlit、Nginx、公网资源、认证下载和生成；记录结果与受保护回退脚本。

部署样例在 `deploy/`，路径/端口须以实际服务为准。健康检查与功能测试、真实上传/邮件/签署分别记录；不把启动成功当作完整安全验收。

## 检查命令

可复现的测试与静态检查：

```powershell
.venv/Scripts/python.exe -W ignore::DeprecationWarning -m unittest discover -p 'test*.py'
npm --prefix frontend test
npm --prefix frontend run lint -- src test
npm --prefix frontend run build
```

本地 MySQL 3307 和私有 runtime 配置存在时，设置进程变量 `PRESENCE_MYSQL_TEMP_TEST=1` 可启用连接级 TEMPORARY 表测试；不创建持久测试员工，不写真实业务表。最新已发布基线为后端 403 项、前端 43 项通过，详见发布记录。
