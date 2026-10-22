# 员工名册第一版

2026-10-22，第一版本地实现及 NIR 专项测试完成，未部署。当前 NIR 验证仅使用合成号码和隔离数据，不录入真实员工社保号。

## 已确认范围

- 独立导航 `Registre du personnel`，入口 `Enregistrer un salarié`；当前向 Admin 与 SuperAdmin 开放。SuperAdmin 继承全部 Admin 功能，普通 Admin 不能授予、撤销或修改 SuperAdmin 权限。本地测试先显示 ChristopheH 的 SuperAdmin 身份，正式环境隐藏角色/账号以后另做。
- 员工登记与 `Créer un compte` 分开。只要求正式姓名，其他字段选填；不会自动创建员工或登录账号，也不从邮箱前缀推断姓名。
- 自动分配固定员工 ID / 登记号，不因排序、姓名或账号变化而重编。
- 基本资料：姓名、国籍、性别、地址、出生日期、职位、资格、人员类型。类型支持 salarié、stagiaire、alternant，也允许尚未填写。
- 社保号支持本地受控录入、替换、清除和显式短时查看；使用 AES-256-GCM 与独立密钥环加密。此前“仅预留、不可录入”的第一阶段说明已被本次实现替代；生产环境默认关闭，尚未发布。
- 登录账号可选关联，一对一，可后补；没有账号或账号停用不影响员工档案。名册归档/登记离职也不改变登录权限。
- 多次入职沿用同一员工身份，新增任职经历。入离职日期、合同类型均选填；CDI、CDD、Convention de stage、Autre，Autre 可补说明。
- 名册先独立维护，不写入旧考勤或合同记录。已有合同只作为手动确认的预填来源；不能将当前合同开始日期直接视为实际首次入职，不能将非 CDI 自动认作 CDD。
- 主列表每人一行，展示姓名、人员类型、职位与最新任职状态；窄屏收起部分列。右侧「…」打开完整档案，查看其他资料、全部任职经历和修改历史。首版网页登记/查询，不提供 Excel 导入或导出。
- Admin 可以更正资料，保留前后快照、操作人、时间，并可归档/恢复；无永久删除权限。后续本地权限升级增加 SuperAdmin 专用永久删除，必须先归档、核对版本并二次确认，详见下节。没有实际删除任何真实记录，尚未部署。

## 页面和接口职责

员工名册的注册、搜索、筛选、刷新、分页和每页条数整合到最上方应用栏。类型、任职状态和归档筛选收在筛选图标内；保留法语标签与键盘操作。工具栏窄屏换行，应用壳通过 ResizeObserver 测量真实高度，表格使用剩余可视空间，不再保留重复大标题或底部分页栏。

列表采用固定列宽分配与响应式列显示，取消原先 2100px 的最小表宽。正文仅纵向滚动；姓名和状态始终可见，完整长文本通过提示或档案查看。每行「…」直接打开全屏档案，不增加一层菜单。

名册的可访问 h1 标题统一由应用栏提供，正文不重复渲染隐藏标题。此前隐藏标题误将 MUI 的数字 `1`（即 `100%`）当作 `1px`，导致最外层溢出；已修正，不用全局隐藏溢出来掩盖该问题。

完整档案在 Informations / Périodes d’emploi / Historique 切页时保持全屏宽高，正文独立滚动，标题与关闭操作不收缩，滚动条预留宽度以避免内容横向跳动。列表收起的国籍、性别、地址、出生日期、资格、账号及任职日期和合同信息仍可在档案中查看。编辑窗口仍保留原布局。

- `frontend/src/personnel/`：请求类型、分页/搜索/筛选表格、注册/编辑/详情、任职经历和修改历史。复用统一请求、法语日期组件、草稿保护和 Light/Dark 主题。`SensitiveNumber.tsx` 管理短时明文显示、取消和自动清除；`useSensitiveNumberCapability.ts` 仅缓存可用性标记。
- `frontend/src/personnel/PersonnelToolbar.tsx`：名册专用应用栏操作与筛选弹层；分页加载期间不传入超出 count 范围的页码。
- `frontend/src/workspaceToolbar.ts` / `components/WorkspaceToolbar.tsx`：应用栏宿主上下文与 Portal；页面持有业务状态，只有当前可见页能渲染顶部操作，隐藏页不遗留按钮或弹层。`AppWorkspace.tsx` 负责宿主和高度测量，八个主页面均已接入；`workspacePages.ts` 维护唯一页面标题及悬停说明。
- `backend/personnel_register.py`：独立路由、权限重查、事务写入及 NIR 专用查看；`personnel_rules.py`：字段、日期、状态及敏感输入屏蔽；`personnel_store.py`：启动时增量建表、持久化、审计和账号更名关联；`personnel_sensitive.py`：NIR 格式规范化、AES-256-GCM 加解密、独立密钥环加载及仅供本地入口调用的密钥初始化。前端隐藏菜单不代替鉴权。
- `backend/local_rebuild.py` / `backend/rebuild_server.py`：显式注册名册路由；本地连接仍固定独立数据副本。新文件未发布前线上无名册入口。
- `backend/main.py`：本人改登录名时，在原事务中同步名册账号关联；不更名员工身份，不改写历史事件中的操作者姓名快照。

接口使用 `/personnel-register` 前缀：列表、新建、`/accounts` 可关联账号查询、`/{id}` 详情/更正、`/{id}/periods` 新增任职、`/{id}/periods/{period_id}` 更正任职、`/{id}/archive` 归档/恢复、`/{id}/history` 分页修改历史。

SuperAdmin 专用 `DELETE /personnel-register/{id}` 请求需 `{expected_version, confirmation}`，confirmation 为精确登记号字符串或正式姓名。服务端锁定用户与员工后重新核对当前角色、会话版本、已归档状态及员工版本；未归档或版本冲突返回 409，确认不符返回 400，普通 Admin 返回 403。删除只作用于该独立名册员工及其任职、含个人资料的本地更正历史和密文列，不删除登录账号、考勤、原合同或 Documents。删除与 `permanent_deletion_events` 全局审计同事务，审计失败全部回滚；全局事件不依附被删员工外键，只留资源 ID、操作者、时间和允许的计数/确认布尔值，不留姓名、地址、生日、NIR 或密文。

NIR 专用接口：`GET /personnel-register/capabilities` 返回 `{nir_enabled: boolean}`；`POST /personnel-register/{id}/social-security-number/reveal` 返回短时明文。两个接口均需实时管理员鉴权。前者是页面能力判断，不是权限凭证。

## 数据与权限

员工是长期身份；账号是可选的登录关系；任职经历独立保存。既有 users、personnel_contracts、Presence 和文件不迁移、不清空，不使用临时测试人员填充列表。

写操作要求传入员工版本，锁定后比较版本，冲突返回 409，不能静默覆盖别人的更正。权限由数据库重新核对当前角色、活跃状态、登录删除标记及会话版本；写入按 users 排序锁、再锁员工，业务变更与审计同事务。账号一对一约束由数据库唯一索引执行。

普通归档仍保留全部资料和历史；SuperAdmin 的明确永久删除会清除该名册中的个人资料历史，保留的仅是独立全局删除元数据审计。这里的删除不声称自动清理离线备份。Documents 与旧合同使用各自独立确认：文件仅 SuperAdmin 可永久删除，必须先归档并确认文件 ID；合同必须先归档并确认合同 ID，任何工作区的保留考勤引用该合同都会拒绝删除，不能级联破坏快照。文件数据库行/删除审计提交后才清理文件；物理清理失败返回 `cleanup_pending`，不能显示为物理删除已经完成。

日期按巴黎当天判断。实际入职日期缺失时显示待补充，不推断在职；未来入职显示待入职；入职已发生且离职为空/未到则为在职，离职日在当天之前则为已离职。该状态指最新有日期的任职段，界面明确此含义；全部任职未填日期时取最近登记段，状态未知。归档是独立维度。列表、账号选择和历史均在服务端分页/限制条数。

存储新增 `personnel_register`、`personnel_register_periods`、`personnel_register_events` 三张空表。关联用户名列的字符集、排序规则沿用既有 users.username，保证唯一性和连接查询一致。元数据 TIMESTAMP 经 MySQL UNIX_TIMESTAMP 读取为 UTC-aware 值，API/快照带明确偏移；不修改数据库全局或连接时区。局部请求验证错误不回显 input/ctx，避免回显被拒绝的个人输入。

本期提供的 Admin 资料更正日志不等同于不可改写的法定电子档案存储。法定名册所需的其他条件性字段、附件、电子留存程序以及正式隐藏超级管理员入口仍属于后续范围。

## 社保号的存储和查看边界

- `personnel_register.social_security_number_encrypted` 是可空独立密文列，启动迁移先检查列再增量添加，不清空已有员工。仅 SQL `NULL` 表示未填写；空白或损坏的密文不是有效清除，查看时失败关闭。
- 公开 Employee DTO、列表、详情和普通历史快照只包含 `has_social_security_number` 布尔值，不包含原号或密文；列表查询不读取密文内容、不执行解密。
- 创建/更新中的 `social_security_number` 为排除普通序列化和 repr 的 `SecretStr`：不传字段保留原值，显式 `null` 或空白清除，非空值规范化后替换。NIR 被关闭或密钥不可用时，NIR 动作返回泛化 503；不带该字段的普通资料更新仍可用，不会清空已有号码。
- 加密使用随机 nonce，认证数据绑定员工不可变 ID、用途、版本与 key ID；密文复制给另一员工、认证失败或缺少旧 key 均不能查看。格式校验允许规范化空格/连字符及支持的科西嘉编码，只检查格式，不验证真实身份、归属或 modulo-97 校验码。
- 更改号码与查看号码分别记为 `social_security_number_changed` / `social_security_number_viewed`。专用事件仅保留员工 ID、是否存在、操作者和时间；替换前后均为“存在”时仍有变更事件，不记录旧号、新号或加密 envelope。普通审计快照也过滤专用原号/密文键。查看必须先成功提交审计，失败不会返回解密结果。
- 所有名册路由响应设置 `Cache-Control: no-store, no-cache, must-revalidate, private`、`Pragma: no-cache`、`Expires: 0`；422 不回显原始 input/ctx，未知字段位置也经过过滤。前端通过直接请求查看，不把返回号码放入 React Query、localStorage 或 sessionStorage；显示 30 秒后清除，关闭、切换员工/标签/工作区、进入编辑或页面隐藏也清除并取消待返回请求。
- 加密范围仅限专用 NIR 字段。姓名、地址、生日等其他资料及其普通审计快照没有因此获得字段级加密；自由文本和附件内手工填写的号码也不会被自动识别或加密。不要把本功能描述为全系统个人数据加密或全面防止误填。

## 独立密钥及本地配置

`PRESENCE_PERSONNEL_NIR_ENABLED` 仅精确为 `1` 时请求开启功能，并且独立密钥环必须有效；未配置时默认关闭。`PRESENCE_PERSONNEL_NIR_KEYS_FILE` 只提供密钥文件路径，不接受或派生登录会话密钥。密钥环 JSON 包含 `active_key_id` 和 `keys` 映射，每项为独立随机 32 字节密钥的 base64url 编码；不在源码、日志、数据库或本文件中填写实际密钥。

本地入口 `backend/local_rebuild.py` 自动在被忽略的 `.local-test/rebuild/keys/personnel-nir.json` 首次创建独立密钥环并启用本地测试；已有文件仅校验、不覆盖。该初始化行为只属于本地入口，服务器入口不会自动生成生产密钥。POSIX 首次创建文件使用 `0600`，目录请求 `0700`；Windows 必须另外依靠 NTFS ACL 限制到运行账号和获授权管理员，不能将 POSIX mode 当作 Windows 权限已经验证的证据。本机本次已配置 NTFS 目录权限，fresh 环境仍需重新配置和检查。

密钥不得随普通数据库快照或源码包备份。需要单独、受限且可恢复的密钥备份；丢失密钥后数据库备份不能恢复原号。轮换时添加新 key ID/材料并切换 `active_key_id`，保留仍被现有密文或历史备份引用的旧 key，不覆盖相同 ID 的旧材料。当前没有后台轮换界面或自动重加密任务；操作前后应使用隔离合成数据验证，生产启用及真实数据使用需另行验收。

## 验证与发布

权限升级补充：`test_resource_deletion.py` 使用隔离 SQLite 和临时合成文件验证名册/文件/合同的 SuperAdmin 专属删除、归档与确认、名册版本冲突、全局审计保留和失败回滚、文件清理失败提示、合同跨工作区引用保护，以及角色撤销/会话失效的事务内重查。未删除真实员工或业务资料；历史 NIR 专项计数见下文，新增权限改动的最终整体验证结果另由发布记录汇总。

已完成 Python AST、TypeScript、生产构建及完整前端 lint；只读代码复核修正了历史动作标签、时间戳偏移及末页归档后分页边界。NIR 专项共 52 项通过：29 项加密/输入规则单测、18 项临时 SQLite HTTP/迁移测试、5 项真实 MySQL 连接级 TEMPORARY 表测试。覆盖加解密、篡改和跨人密文、保留/替换/清除、越权、失效会话、版本冲突、缺密钥/开关、审计失败回滚、校验/日志脱敏及无缓存响应。

全量 Python unittest discover 共 237 项：231 项通过、6 项按 opt-in 配置跳过（5 项新增 MySQL TEMPORARY 测试、1 项原有 MySQL 测试）；上述新增 MySQL 5 项已另行 opt-in 运行通过。前端 TypeScript、完整 lint、7 项既有 draftRules 测试及 production build 均通过。此处的专项计数与全量回归存在包含关系，不应相加作为独立测试总数。

本地 API 8013 已重启并完成启动迁移，本地独立密钥就绪；未认证的 capabilities 请求返回 401 且禁止缓存。浏览器工具拒绝读取 http://127.0.0.1:5173/（URL 策略），因此未确认浏览器视觉和输入交互，仍需用户本地手动确认。测试仅使用合成号码、临时文件与隔离测试表，没有向真实员工档案写入号码，没有真实邮件或签署，没有上线。SQLite 事务模拟及单连接 MySQL TEMPORARY 测试不等同于生产负载或多连接并发验收。

本地启动脚本将 MySQL 临时目录固定到被忽略的 `.local-test/rebuild/mysql-tmp`，使用指定 loopback 端口探测并等待数据库就绪，避免受限环境无法列举系统 TCP 状态、C:\\tmp 不可写及恢复时间超过两秒造成的 API 提前启动。Vite 子进程被沙箱 EPERM 阻止后，已通过执行审批在沙箱外完成构建和本地启动。

发布前须按 backend/README.md 备份，并比较所有可访问 API 的权限及关联更名兼容性。正式上线前应再进行真实人员手动验收与隔离测试。
