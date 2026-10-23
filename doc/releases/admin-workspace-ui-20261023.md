# 2026-10-23 Admin 导航分组上线

- 前端：`admin-workspace-ui-20261023T151905Z`，功能提交 `63ef6c3935f2`。
- 员工名册和 Administration 集中在带边框的 Espace Admin 分组，仅 Admin/SuperAdmin 可见；进入相应页面时顶栏显示区域标识。普通功能独立排列。
- TypeScript / Vite 生产构建与 `git diff --check` 通过。本次未新增或执行自动化测试。
- 发布前原版本 297 个文件摘要一致；发布后 247 个静态文件校验通过，公网首页及 9 项资源与发布包一致，3 个私有接口拒绝匿名访问。
- 四项服务 active，两套 API 的进程与原版本 297 个文件均未改变。此次只切换 Nginx 静态目录；账号、数据库、密钥和业务文件不参与部署。未发送邮件或签署请求。
- 新目录：`/var/www/presence-react/releases/admin-workspace-ui-20261023T151905Z/frontend/dist`。保留旧 assets，兼容已打开的页面。
- 备份：`/var/backups/presence-react/admin-workspace-ui-20261023T151905Z`，含旧前端、切换前后 Nginx 配置、版本清单与发布核对结果。

## 本次界面回退

```sh
sudo /var/backups/presence-react/admin-workspace-ui-20261023T151905Z/rollback-ui.sh
```

脚本核对当前界面版本后恢复上个前端并 reload Nginx；不改 API 或数据库。语法检查通过，未实际执行线上回退。

API 仍为 `account-names-20261023T150121Z`。若需执行[上一版协调回退](account-names-20261023.md)，先回退本次独立前端，以满足上一版配置保护检查。
