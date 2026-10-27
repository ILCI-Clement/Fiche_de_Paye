# Responsable 独立浏览器测试（2026-10-27）

测试页面：`http://127.0.0.1:5174/`，API 仅监听 `127.0.0.1:8014`。原 5173 / 8013 测试副本保持独立。

使用本机 MySQL 3307 的 `presence_responsable_qa_20261027`，从本地副本仅复制表结构，没有复制真实业务数据。初始化、进程日志、结果和凭据位于 Git 忽略的 `.local-test/responsable-qa/`；`access.json`、`runtime.json` 已限制为本机当前用户可读写，密码不进入源码、说明或命令输出。不要将该目录打包部署。

| 账号 | 用途 |
| --- | --- |
| `responsable.qa@example.com` | 唯一角色 Responsable；管理 TEST 部门 |
| `employee.qa@example.com` | 合成直接下属，关联合成名册 |
| `outsider.qa@example.com` | 其他团队，验证越权拦截 |

另有不可供日常登录的初始化 Admin；其随机密码未保存。账号由现有 `create_user` 共用逻辑创建并记录角色审计。QA 入口关闭建号、邮件、合同提醒和签署发送，无生产密钥；全部服务绑定 loopback。

## 已完成

- 用新建 Responsable 实际登录浏览器，侧栏无 Admin 区域，选择器仅显示直接下属；一条未关联历史表未阻塞页面且不泄露内容。
- UI 中补录合成合同（2026-10-01 至 2026-10-31），自动创建工时表。
- UI 将 10 月 1 日上午设为 Congé payé 并应用，刷新后仍保存；没有模拟浏览器令牌或跳过认证。
- 实际 HTTP 登录验证角色只有 Responsable；越权打开其他团队、读取名册、修改已有合同均被拒绝；重复补录返回 409；普通员工无考勤编辑权限。
- SQL 只读核对隐藏历史逐字段一致、合同仅一份、补录审计仅一条，名册同步合同日期且实际离职保持空。

执行结果：`.local-test/responsable-qa/results.json`；浏览器截图：`browser-result.png`。本次仅增加隔离测试数据，没有部署或修改原测试副本中的真实合同。
