# 2026-10-23 正式版签署列表简化

- 前端版本：`production-signatures-ui-20261023T120519Z`，功能提交 `5caa88c`，已推送 `origin/main`。
- 正式版隐藏 Production / Tests SPECIMEN 切换及测试说明，列表和月度导出固定使用 production。本地副本保留切换；环境未就绪时也默认正式模式。
- 考勤与 Profile 共用签署列表，因此同时生效。没有删除测试文件或改动后端权限。
- TypeScript、lint、生产构建通过；公网页面、RH 图标及新签署组件资源哈希匹配，环境接口为 server，匿名文件列表拒绝访问，回调 GET 为 405，四项服务均 active。本次未新增签署或发送邮件。
- 发布前核对原静态资源 247 项一致，备份 Nginx 与原 dist；保留旧哈希资源，Nginx 检查后重载。
- 新静态目录：`/var/www/presence-react/releases/production-signatures-ui-20261023T120519Z/frontend/dist`。API current 链接仍指向 `signed-archive-20261023T111516Z`，没有重启 API 或修改数据库、密钥及业务文件。
- 仅 root 可读的界面备份：`/var/backups/presence-react/production-signatures-ui-20261023T120519Z`，含配置、静态文件、发布清单及验证记录。

回退本次界面调整：

```sh
sudo /var/backups/presence-react/production-signatures-ui-20261023T120519Z/rollback-ui.sh
```

脚本先检查当前界面版本，再恢复前次签署归档界面；API、回调 worker 和所有业务数据保持当前状态。
