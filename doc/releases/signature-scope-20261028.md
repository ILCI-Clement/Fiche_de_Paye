# 2026-10-28 签署权限与组织保存修复发布

- 版本：`scope-access-20261028T093201Z`；代码提交：`08e005a04d673bb458f6a45b8e32535e17983723`。
- 正式地址：https://app.presence.217.65.145.103.nip.io/ 。React API 8015、保留 API 8001 和前端同时更新。
- 发布前已比较服务器源码；与上一版发布清单一致，无未合并的服务器修改。

## 行为

签署发送前保存当次实际签署负责人。Responsable 查看列表、Profile、预览和下载，必须同时是员工当前直接负责人和当次签署负责人；Admin/SuperAdmin 和员工本人保留完整历史。旧记录归属不根据发送操作者或当前组织猜测，未知归属仅对负责人隐藏，不删除文件。

同值保存部门不再错误返回“Groupe introuvable”。管理部门必须配合显式 Responsable 角色；界面提供对应校验。组织保存失败原因显示于确认弹窗内部，成功后关闭弹窗。

## 验证

- 本地后端 436 项、前端 52 项通过；TypeScript、lint、生产构建通过。
- 隔离 loopback 候选 API：141 项基础检查及 35 项签署权限检查。
- 正式两套 API：233 项基础检查及 69 项签署权限检查，包含 48 次直接文件越权请求被拒绝。
- 11 个公开页面/静态资源哈希匹配；两套 API、旧界面和 Nginx 均 active，新 API 启动日志无 traceback。
- 35 张原有表的原有列数据核对通过；仅追加可空 `signed_fiche_documents.manager_username`，历史归属保持 NULL。
- 未导入本地测试账号或数据，未修改正式账号角色/密码；发布测试未发送邮件或创建签署请求。

## 备份与回退

受保护备份目录：`/var/backups/presence-react/scope-access-20261028T093201Z`，包含数据库、源码/配置、私有文件、签署 SQLite、核对结果及回退脚本。凭据与密钥仅留于服务器受保护备份中。

界面回退保留新的后端权限：

```sh
sudo /var/backups/presence-react/scope-access-20261028T093201Z/rollback-ui.sh
```

协调回退恢复上一版前后端及服务配置，不自动还原数据库或移除新增列；也会恢复上一版较宽的签署访问规则。只有确认当前指针仍属于本版才执行：

```sh
sudo /var/backups/presence-react/scope-access-20261028T093201Z/rollback.sh
```
