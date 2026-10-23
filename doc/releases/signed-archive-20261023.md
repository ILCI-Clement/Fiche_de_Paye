# 2026-10-23 已签考勤归档发布

- 版本：`signed-archive-20261023T111516Z`。
- 功能代码提交：`fa319fb`，已推送 `origin/main`。
- 上一版本：`personnel-polish-20261022T150919Z` / `0b2775f`。
- 线上：<https://app.presence.217.65.145.103.nip.io/>。

## 功能

- 双方全部签署后接收 ClawShow 通知，持久排队下载最终 PDF，按考勤年/月保存原始字节。
- 考勤顶部新增 Signatures；个人 Profile 可查看月度文件。眼睛预览、独立下载、Admin 月度 ZIP，20 项分页；SPECIMEN 与正式文件分开。
- 本人、当前直接负责人及管理员由后端鉴权；月度全量 ZIP 和下载重试仅管理员。重复通知去重，失败退避重试，租约协调新旧 API，文件不覆盖并检查 SHA-256。
- 同时发布已完成的 ILCI · Espace RH 更名、社保号表格行内短时脱敏查看、密码连续错误 5 次冷却 10 分钟及未满 5 次闲置 10 分钟重置。
- 保持原生产角色、密码、会话签名密钥、NIR keyring 与数据。没有提升生产 Super Admin，没有导入本地副本或测试账号。

## 测试与核对

| 检查 | 结果 |
| --- | --- |
| 后端完整回归 | 403 项通过，启用本地连接级 MySQL TEMPORARY 表集成测试 |
| 前端 | 43 项通过，TypeScript、lint、Vite 生产构建通过 |
| 用户手动沙盒双签 | 1 份已通过实际回调回存；原文件大小、SHA-256、PDF 结构正确，读取/下载有事件记录 |
| 新功能测试 | 回调伪造、先后顺序/重复通知、严格完成标志、URL 限制、租约过期/旧 worker、下载失败重试、损坏文件、ZIP、分页、当前负责人/会话版本、账号改名、登录冷却及别名计数 |
| 线上候选 API | 108 项通过 |
| 发布后新旧 API | 176 项通过 |
| 公网 | 11 项页面/资产匹配，另核对 RH favicon；匿名私有接口拒绝；回调 GET 为 405、缺少凭据的 POST 为 404 |
| 服务 | presence-react-api、presence-app、presence-interface、nginx 均 active |
| 原有数据 | 发布前 32 张表原列与原数据核对一致；新增登录保护列及两张归档表，不清空历史 |

初轮发现旧测试夹具缺少新增字段/表，已补齐。人员解析器恢复默认只返回显示字段，归档调用显式获取内部身份，避免改变已有接口显示约定。测试使用独立临时文件/表，没有向生产发送测试签署或通知邮件。

已存在一份负责人考勤工作区因历史资料超出当前管理范围返回 409；与发布前行为一致，保留该防护，没有删除历史资料来规避。

## 生产配置与备份

- 程序：`/var/www/presence-react/releases/signed-archive-20261023T111516Z`；current 已指向该目录。
- 生产回调：网站 `/api/esign/callback`，两套 API 的 `70-signed-archive.conf` 显式配置。
- 永久目录：`/var/lib/presence-app/signed-fiches`，服务身份 `streamlitapp`、0700；必须和 MariaDB 归档元数据一起备份。
- 两套 API 下载 worker 启用、数据库租约协调；生产原有 `sk_live_` 凭据保留，没有复制本地 Sandbox key。
- 完整备份：`/var/backups/presence-react/signed-archive-20261023T111516Z`，包含数据库 SQL、源码/配置/NIR keyring、私有文件、签署 SQLite 和数据核对基线，均仅 root 可读。
- 保留旧哈希静态资产和既有 Python 环境；未触碰其他项目的 8002 服务。

## 回退

```sh
sudo /var/backups/presence-react/signed-archive-20261023T111516Z/rollback-ui.sh
```

只恢复上一版界面，保留当前 API 的回调接收、归档 worker、登录冷却及账号保护。不要直接将后端回退到不支持回调的旧版本，也不要恢复数据库来覆盖发布后的新签署和业务记录。完整代码/配置及数据库备份用于单独审核的数据恢复。

## 范围

旧请求未携带本版回调凭据，不能自动补收，也不会自动重新发送邀请。只有发布后创建的新请求自动归档。已验证真实 Sandbox 回调；没有新增 Production 签署来消耗正式配额。PDF 原字节保留与哈希核对不等于独立验证证书信任链。
