# Presence 后端

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

- 版本：`personnel-polish-20261022T150919Z`；Git 提交见服务器 `git-release.json`。
- 应用链接：`/var/www/presence-react/current`。
- 实际目录：`/var/www/presence-react/releases/personnel-polish-20261022T150919Z`。
- root 专用备份：`/var/backups/presence-react/personnel-polish-20261022T150919Z`。
- 完整验证与已知问题：[发布记录](../doc/releases/personnel-polish-20261022.md)。
- 社保号配置：已单独启用生产加密录入/查看，仅管理员可用；[密钥与开关记录](../doc/releases/nir-enable-20261022.md)。应用源码仍为上述版本。

安全回退界面：

```sh
sudo /var/backups/presence-react/personnel-polish-20261022T150919Z/rollback-ui.sh
```

脚本恢复上一版静态界面，保留新版 API 的账号状态/会话版本保护及业务数据。不要直接运行不识别 `is_active`、`session_version`、`login_deleted_at` 的旧 API；它可能重新放行停用或删除登录的身份。数据库恢复须单独决定，不能覆盖发布后的业务写入。

旧发布记录见 [历史时间线](../doc/releases/20261014-20261021.md)。此次代码清理尚未发布。

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
| `CLAWSHOW_ESIGN_API_KEY`、`CLAWSHOW_ESIGN_NAMESPACE` | 服务商凭据/命名空间，默认命名空间 `ilci` |
| `PRESENCE_ESIGN_ORGANIZATION_NAME`、`PRESENCE_ESIGN_ORGANIZATION_FULL_NAME` | 签署邮件机构名称 |
| `PRESENCE_ESIGN_REPRESENTATIVE_NAME`、`PRESENCE_ESIGN_REPRESENTATIVE_TITLE` | 邮件落款；默认解析负责人和 Responsable |
| `PRESENCE_SIGNATURE_ATTEMPTS_DB` | 持久签署回执 SQLite；当前 `/var/lib/presence-app/react/signature-attempts.sqlite3` |
| `PRESENCE_CONTRACT_REMINDER_WORKER` | React 设为 0，由保留 API 运行唯一自动提醒任务；停用旧服务前须移交任务 |
| `PRESENCE_PERSONNEL_NIR_ENABLED`、`PRESENCE_PERSONNEL_NIR_KEYS_FILE` | 生产 NIR 默认关闭；仅精确开关 1 且独立 keyring 有效时开放 |

依赖见 `requirements.txt`；PDF 转换还需 LibreOffice。AES-256-GCM 使用独立密钥，不从会话密钥派生。POSIX keyring 使用 0600；Windows 须另设 NTFS ACL，密钥独立备份且轮换时保留仍被引用的旧 key。

## 数据与关键行为

- 登录接受绑定邮箱或用户名，大小写及标识两端空白不敏感；多身份匹配统一拒绝，密码空白保持原样。
- 账号停用/角色变更递增会话版本并撤销持久登录；删除登录保留身份及业务历史且不可重新启用。物理删除仅 SuperAdmin，先停用、精确确认并处理依赖，不级联删除他人文件。
- 员工名册与登录账号、考勤合同独立。Admin 手动一对一关联；本人只读，直接负责人不因此获得名册隐私。关联版本检查、解除/替换和审计同事务提交，失败回滚。
- `/account-personnel/{username}` 仅本人/Admin；`/choices`、`/link` 仅 Admin。关联提交包含 `employee_id`、`expected_employee_id`、`expected_version`、`target_version`，无关联值为 null。
- 人名按显式关联的名册姓名显示，否则使用有效邮箱前缀；身份和权限仍以账号 ID 判断，历史快照不自动回写。
- 考勤保存按版本和稳定 ID 检查；历史表不因当前合同更改而丢失。新建须有有效合同；Admin 用归档，SuperAdmin 删除还需明确 `deleted_ids`。
- Documents 接受单份 PDF/JPG/PNG，最大 10 MiB。员工/经理只能为本人上传，Admin 可代传；文件私有下载，ZIP 最多 100 份/100 MiB，失败不返回部分结果。
- 文档编号按巴黎上传日分配 `YYYYMMDD-0001`；行锁保证每日序号唯一，历史缺号启动时补齐，归档/改月不改号，删除不回收。
- 签署由员工先签、当前直接负责人后签。LibreOffice 转换后按最终 PDF 定位签名，拒绝不明确位置；临时公开 PDF 默认保留 24 小时。签署接受回执不代表完成，目前无最终已签 PDF 自动回存。
- 请假按半天申请，审批前可撤回，跟随当前直接负责人或 Admin 审批并留历史；不自动计算余额或写考勤。
- `GET /ui-capabilities` 只返回环境和布尔能力；就绪检查不证明实际邮件投递、外部 PDF 获取或业务权限。

NIR 原号与密文不进入普通 Profile；专用查看须先提交审计，禁止缓存，前端短时显示后清除。更多字段和接口规则见 [员工名册说明](../doc/personnel-register.md)。

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

本地 MySQL 3307 和私有 runtime 配置存在时，设置进程变量 `PRESENCE_MYSQL_TEMP_TEST=1` 可启用连接级 TEMPORARY 表测试；不创建持久测试员工，不写真实业务表。最新已发布基线为后端 376 项、前端 43 项通过，详见发布记录。
