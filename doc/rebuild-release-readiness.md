# 发布与回退索引

## 当前线上版本

- 版本：`personnel-polish-20261022T150919Z`；源码提交见服务器 `git-release.json`。
- 网址：https://app.presence.217.65.145.103.nip.io/ 。
- React API：`presence-react-api`，127.0.0.1:8015。
- 保留服务：`presence-app`（8001）和 `presence-interface`（8501）。
- 当前代码：`/var/www/presence-react/current` → `releases/personnel-polish-20261022T150919Z`。

完整范围、检查结果、备份内容和已知问题见 [最新发布记录](releases/personnel-polish-20261022.md)。

## 安全回退

```sh
sudo /var/backups/presence-react/personnel-polish-20261022T150919Z/rollback-ui.sh
```

恢复上一版静态 UI，保留新 API 的账号停用、会话版本和删除保护；不恢复数据库、不删除上传文件。不要直接恢复不识别账号状态的旧后端。

备份目录仅 root 可读：`/var/backups/presence-react/personnel-polish-20261022T150919Z`。源码、数据库、私有文件和签署回执 SQLite 均已备份。后端或数据库恢复须先制定兼容与数据保留方案。

## 当前验收边界

- 已发布基线：后端 376 项（含本地 MySQL 临时表）、前端 43 项通过；146 次发布后 API 检查，10 个公网资源字节匹配，认证下载与 Excel/PDF 生成正常。
- 仍有 1 个负责人历史考勤工作区因包含当前权限外记录而返回 409；旧版已有相同拦截，需 Admin 核对，不自动删改历史。
- 没有重新发送真实邮件/签署，也没有完整浏览器端到端或并发负载验收。
- NIR 已以独立生产密钥显式开启，仅管理员录入/查看；[启用与关闭记录](releases/nir-enable-20261022.md)。不导入本地数据、密钥或 SuperAdmin 测试赋权。
- 登录限流、OTP/CAPTCHA、邮箱验证、Microsoft 邮件集成、最终签署自动归档、可配置文档分类仍需后续处理。

## 发布前阅读

1. [后端运行与部署](../backend/README.md)：配置、依赖、备份及两套 API 一致性。
2. [模块索引](modules.md)：本次改动涉及的入口与直接依赖。
3. [历史发布记录](releases/20261014-20261021.md)：旧版本、事件和原回退路径，仅供追溯。
