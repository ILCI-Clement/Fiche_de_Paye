# Documents 自定义类型（2026-10-28 已上线）

每份文件使用一个类型。Admin / SuperAdmin 在 Documents 顶部点击 **Gérer les types**，可以新增、改名、停用或恢复自定义类型。普通员工和 Responsable 只能选择有效类型，不能维护类型或更正文件分类；既有文件查看、代上传、归档及删除权限保持原规则。

分类标签由 `DocumentTypeBadge.tsx` 统一展示：圆角 1px 彩色描边与浅色填充；上传/分类/筛选选项、文件行及类型管理使用相同配色。颜色按稳定代码分配，改名不变色，兼容浅色和深色主题；停用项在选择和管理界面使用灰色虚线框。

员工点击 **Déposer un document** 后从 **Type de document** 中选择；Admin 可从已有文件的 **… → Modifier le type** 更正分类。筛选框列出全部类型（包括已停用项）以查询历史；搜索和 ZIP 文件名使用可读分类名称。改名沿用原分类代码，历史文件随之显示新名称；停用只限制新选择，保留旧文件及标签。Transport / Autre 为兼容旧页面固定保留，Autre 仍可补充说明，未分类的旧文件不会自动归类。

## 存储与接口

- 新增 `document_categories` 和只追加的 `document_category_events`；代码不超过原有 `transport_receipts.document_type VARCHAR(20)` 容量，无需更改发票表或文件。
- 启动迁移使用数据库命名锁，幂等初始化两个标准类型，不重写旧发票。自定义类型使用随机稳定代码；分类名称限定 1–80 个可见字符，按大小写、重音和空白归一化去重，已停用的名称仍被保留。
- `GET /document-categories`：已登录用户读取，返回类型列表和实时 `can_manage`，禁止响应缓存。
- `POST /document-categories` / `PATCH /document-categories/{code}`：数据库内重新核对有效 Admin / SuperAdmin，会话及角色不能由客户端覆盖。更新必须匹配 `expected_version`，冲突 409。
- 上传接口在同一事务中锁定并检查有效类型；类型不存在 400、已停用 409，不留下文件或发票行。
- `PATCH /transport-receipts/{id}/category`：仅 Admin / SuperAdmin，提交原 `expected_document_type`、`expected_other_type` 防止并发覆盖；只更改分类元数据，文件内容、归属、编号、日期和归档状态不变。
- 类型变更与发票重新分类记录操作者及修改前后内容；审计失败整笔回滚。无永久删除类型接口。

## 验证与发布

470 项后端回归（含真实 MySQL 临时表）、54 项前端测试、TypeScript、lint、生产构建通过。专项测试 `test_document_categories.py` 覆盖权限、停用、去重、并发版本、审计回滚、文件不变与旧类型兼容；`test_document_categories_mysql.py` 验证重复迁移不改旧记录、20 字符代码及数据库约束。

本地 5174 / 8014 隔离 QA 使用合成 Admin / 员工、新分类 Hôtel 和合成 PDF 验证管理、上传选择、标签展示及更正入口；未改动生产账号/文件，未发送邮件或签署。

尚未部署。发布需前后端一起更新，flat 保留 API 同步 `documents.py` 与新增 `document_categories.py`，先按 backend/README.md 比对和备份。没有新增密钥或环境变量；新分类表和审计表应随业务数据库一起备份。
