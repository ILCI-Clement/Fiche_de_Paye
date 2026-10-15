# Congés payés：新前端试点

## 状态与范围

2026-10-14：第一版代码和本地预览已完成，尚未发布到线上。此试点独立于 Streamlit 界面。正式接口接入现有认证和 MariaDB；默认本地预览使用虚构账户与独立 SQLite 数据库。

已执行 TypeScript 类型检查、生产构建、Oxlint 静态检查和 Python 编译检查。浏览器已查看员工列表与申请窗口。尚未执行完整的提交、撤回、审批、并发和 MariaDB 回归测试；这些检查是发布前的后续步骤。此前项目的 78 项回归结果属于上一版，不能作为此新模块的验证结果。

## 已确认的业务规则

- 所有已登录账户都可为自己提交带薪假申请。日期可在过去，允许重叠，暂不计算余额或天数，不排除周末或假日。
- 支持全天、上午、下午和跨日期区间；中间日期为整天。单日申请只使用一种时段，开始、结束时段自动联动。
- 员工备注、审批意见均选填，最多 2000 字符。
- 当前直接负责人处理申请；负责人改变后，未完成申请立即跟随新关系。部门负责人权限不替代直接负责人。
- 没有负责人时由 Admin 处理。Admin 可查看全部记录、代审及审批自己的申请。
- 普通负责人不能审批自己的申请，自己的申请交给上级或 Admin。
- 员工只能撤回自己的待审批申请。已经批准或拒绝的申请不能再次处理或撤回。
- 员工查看个人历史，负责人查看当前直接下属历史；撤回记录不出现在普通列表，数据库保留其历史。
- 第一版只有系统内待审批数量、待处理列表及 30 秒刷新，没有邮件通知、考勤自动写入和假期余额。

## 技术栈与入口

新前端支持 Light / Dark，默认深色；登录页和登录后的右上角均提供太阳/月亮按钮。选择保存在本浏览器的 `localStorage`，刷新、退出或重新登录后沿用。配色沿用原 Streamlit：红色强调色与中性灰背景，深色同步覆盖表单、弹窗、标签和列表。主题由 `ThemeModeProvider.tsx` 管理，`theme.ts` 集中提供配色，`components/ThemeToggle.tsx` 复用切换按钮。

| 层 | 技术 | 总入口 |
| --- | --- | --- |
| 新前端 | React、TypeScript、Vite、Material UI、TanStack Query | `frontend/index.html` → `src/main.tsx` → `src/App.tsx` |
| 正式后端 | 现有 FastAPI、PyMySQL、MariaDB | `backend/main.py` 注册 `leave_requests.py` |
| 独立预览后端 | 同一业务模块、SQLite、虚构身份选择 | `backend/leave_preview.py` |

选择 React SPA 是为了逐步迁移现有登录工具，复用 Python 业务层。Material UI 提供表单、弹窗和响应式控件；页面使用 CSS 网格，电脑侧栏在窄屏收起。TanStack Query 管理 API 缓存、加载、刷新及操作后的数据更新。

## 脚本职责

| 文件 | 职责 |
| --- | --- |
| `frontend/src/main.tsx`、`theme.ts` | 挂载 React、统一主题、API 查询缓存 |
| `frontend/src/App.tsx` | 登录状态、导航、个人/团队列表、过滤、分页、待审批数量 |
| `frontend/src/api.ts`、`types.ts` | 唯一 HTTP 请求入口、错误处理、TypeScript 数据契约 |
| `frontend/src/leaveFormat.ts` | 日期、时段、状态名称的法语显示 |
| `frontend/src/components/Login.tsx` | 正式账户登录或独立演示账户选择 |
| `frontend/src/components/LeaveForm.tsx` | 本人的日期区间、半天和选填备注 |
| `frontend/src/components/LeaveDetail.tsx` | 详情、操作历史、审批及撤回确认 |
| `frontend/src/components/StatusChip.tsx` | 状态标签 |
| `frontend/src/index.css` | 电脑、平板和手机布局，无 Streamlit CSS 覆盖 |
| `backend/leave_rules.py` | 输入模型、日期/备注校验、直接负责人和 Admin 权限 |
| `backend/leave_store.py` | 数据库事务、建表、查询、历史写入、账号改名及注销关联 |
| `backend/leave_requests.py` | HTTP 接口编排、按当前身份过滤、状态转换，不反向导入入口 |
| `backend/leave_preview.py` | 仅本地使用的虚构账户与数据库连接，不能部署到生产 |

## 本地打开

在仓库根目录打开一个 PowerShell：

```powershell
python -m uvicorn backend.leave_preview:app --host 127.0.0.1 --port 8012
```

另一个 PowerShell：

```powershell
cd frontend
npm ci
npm run dev
```

浏览器打开 `http://127.0.0.1:5173/`。当前已启动的实例可以直接访问。演示账户：Alice / Fabrice 的负责人是 Robin；Robin 的负责人是 William；William 为 Admin；另有无负责人账户。演示账户均为虚构身份，不读取生产人员信息。

SQLite 文件为 `.local-test/leave-preview.sqlite3`。申请和操作历史在服务重启后仍保留；演示登录令牌只保留在进程内，后端重启后重新选账户。此数据库与原考勤、凭证存储完全分开。

## 手动体验顺序

1. 选择 Alice，点击 `Nouvelle demande`，选日期、全天或半天，填写可选备注，提交。
2. 查看 Alice 的 `Mes demandes`，打开 `Détails` 查看提交时间和当前负责人。
3. 退出，选择 Robin，点击 `Mon équipe` 或通知图标，在待审批列表打开 `Traiter`。
4. 可选填写意见，点击 `Approuver` / `Refuser` 并确认。再用 Alice 查看结果和真实操作人。
5. 另建一份待审批申请，使用 Alice 的 `Retirer ma demande` 并确认，普通列表将隐藏它。
6. 使用无负责人演示账户提交，切换 William 查看 `Toutes les demandes` 并处理。

以上是用户可执行的体验步骤，不表示这些流程已经完成回归验证。

## 接口与存储

| 方法 | 路径 | 行为 |
| --- | --- | --- |
| GET | `/leave/session` | 当前账户、权限和当前负责人基础显示 |
| GET | `/leave-requests?scope=mine\|team&status=all\|pending\|approved\|rejected&page=1` | 分页列表及状态数量，后端确定权限范围 |
| POST | `/leave-requests` | 仅为当前账户创建，拒绝额外身份字段 |
| GET | `/leave-requests/{id}` | 授权范围内的详情和历史 |
| POST | `/leave-requests/{id}/withdraw` | 员工撤回自己的待审批申请 |
| POST | `/leave-requests/{id}/decision` | 当前负责人或 Admin 批准/拒绝 |

正式后端启动时新建 `leave_requests` 和 `leave_events`，不迁移或清空原业务数据。申请内容和操作记录位于同一数据库事务；状态更新只接受 `pending`，同时处理冲突返回 409。MySQL/MariaDB 使用行锁，SQLite 预览使用写事务。

历史记录包含操作人账号和名称快照、UTC 时间、旧/新状态、意见。账号改名同步申请所属账号，但不改历史操作人快照。永久删除账号时解绑申请所有者，保留申请及事件供 Admin 查看，避免同名新账户继承旧申请。没有对外暴露编辑或删除操作历史的接口。

## 正式账户接入与发布准备

正式 `backend/main.py` 使用现有 `/login` 和 Bearer 会话认证。新前端通过同源 `/api` 转发，开发时由 Vite 代理，生产由反向代理负责。只向浏览器提供必要的身份和申请信息，不下发服务器 API key。

如需本地前端连接已升级的真实后端，在 `frontend/.env.local` 覆盖：

```dotenv
VITE_PREVIEW=false
PRESENCE_API_PROXY=http://127.0.0.1:8001
```

后端须已包含本次四个正式模块；现有线上后端尚未发布这些接口，不能直接使用新申请功能。独立 `leave_preview.py` 不随生产发布，生产构建也不显示演示登录。真实登录会修改真实申请，请在明确的测试账户下验证。

发布前按 `backend/README.md` 比较本地与线上差异并备份源码和 MariaDB。同步入口及 `leave_rules.py`、`leave_store.py`、`leave_requests.py`，配置前端静态站点和 `/api/` 代理。至少验证身份隔离、当前负责人变更、Admin 权限、状态竞争、历史保存、账号改名/删除、手机布局和服务重启后的数据持久性，再切换正式使用。

当前构建产物为 `frontend/dist/`，未纳入 Git。依赖精确版本保留在 `frontend/package-lock.json`；开发要求 Node.js 22.12+ 或 24 LTS。
