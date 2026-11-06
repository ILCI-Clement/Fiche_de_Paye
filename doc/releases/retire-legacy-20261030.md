# 单一 RH 服务：2026-10-30

执行时间：2026-10-30 17:12（Europe/Paris）。应用版本仍为 `rh-email-identity-20261030T155220Z`，本次为基础设施收敛，不修改业务代码或资料。

## 当前状态

- 唯一后端为 `presence-react-api`，运行 `backend.rebuild_server:app`，监听 loopback 8015；前端由 Nginx 提供 React 静态文件。
- `presence-app`（8001）和 `presence-interface`（8501）均 stopped、disabled 并 masked；旧启动文件及 drop-in 已移入受保护备份，不再驻留运行或开机启动。
- `/var/www/presence-app`、`/var/www/presence-interface` 旧执行目录移入备份的 `retired-runtime/`。当前 Python 环境不依赖这些目录；其他项目使用的共享依赖未卸载。
- 新增 `/etc/systemd/system/presence-react-api.service.d/zz-single-api.conf`，将合同提醒迁移到新版，统一启用提醒、签署回存、员工快照三个 worker。候选服务仍应关闭 worker。
- 旧 API 域名 `presence.217.65.145.103.nip.io` 只作为新版 8015 的代理别名，保留原路径供历史 PDF 链接、签署回调兼容；不再有单独旧 API 进程。
- `/etc/presence-app`、`/var/lib/presence-app` 和 `/var/lib/presence-secrets` 是当前仍使用的配置、数据及密钥目录，必须保留。

## 验证

执行前对源码发布清单、Nginx 路由、运行进程、服务配置、端口和旧目录依赖进行核对。只有两个被下线进程依赖旧目录，当前服务不依赖。

- 本地入口测试通过，执行脚本语法校验通过。
- 下线后新版 API 367 项主检查通过；邮箱专项 25 项、签署权限 78 项、账户/签署预检 13 项通过，其中拒绝 62 次越权文件访问。
- 8001、8501 无监听；旧单元均 masked；Nginx 不再代理这两个端口。
- 新版服务和 Nginx active，新启动日志无 traceback，生产 worker 配置均启用。
- 公网页面与现有发布文件一致；两个入口均指向新版 API，未认证访问和回调方法检查正确。
- 原有 41 张业务表全部保持一致。账号、密码、角色、合同、请假及签署记录未被此操作修改。未显式发送测试邮件或签署请求，日常自动提醒恢复由唯一服务按已有设置运行。

## 备份及恢复边界

受保护目录：`/var/backups/presence-react/retire-legacy-20261030T161202Z`。

包含数据库 dump、私有文件、源码及配置归档、密钥材料、签署 SQLite 在线备份、数据哈希检查点、验证结果和退役的执行目录及 systemd 文件。备份仅 root 可访问，不上传仓库。

正常发布及回退都应维持单一服务，不再使用会重启旧服务的历史双 API 回退脚本。灾难恢复时先核对当前业务写入和邮箱复用状态；禁止直接恢复旧数据库、遗漏加密密钥或同时启动两份合同提醒 worker。`backend/main.py` 仍是当前入口共享核心，不能作为“旧 API”源码删除。
