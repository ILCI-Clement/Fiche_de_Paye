# 脚本职责与应用入口

当前入口（2026-10-16）：`frontend/index.html` → `frontend/src/main.tsx` → `frontend/src/App.tsx` → `AppWorkspace.tsx`。React 已于 2026-10-15 上线，线上后端使用 `backend/rebuild_server.py`，本地真实数据副本使用 `backend/local_rebuild.py`。登录统一使用真实账号，虚构身份预览已删除。当前模块职责见 [模块索引](modules.md)，启动方式见 [前端说明](../frontend/README.md)。

以下第 1–4 节记录旧 Streamlit 的调用关系，仅用于维护回退版本；其中仍被新版复用的生成、权限及数据兼容模块继续保留。

本文解释程序怎样启动、脚本如何协作，以及本次 Mon équipe 拆分。按功能查文件时，先读 [模块索引](modules.md)。

## 1. 两个总入口

| 入口 | 执行方式 | 负责什么 |
| --- | --- | --- |
| `online_test.py` | `python -m streamlit run online_test.py` | 页面配置、登录恢复、角色导航、执行选中的页面。名称保留历史兼容，现为旧版回退前端入口 |
| `backend/main.py` | 本地包导入 `backend.main`；服务器 `uvicorn main:app` | 创建 FastAPI、读取环境配置、连接数据库、认证和组织权限、注册功能接口及启动迁移 |

前端不直接连接数据库。前端通过 `api_client.py` 携带登录令牌请求后端，最终访问权限由后端判断。生产前端目录为 `/var/www/presence-interface`，后端目录为 `/var/www/presence-app`。密钥保留在服务器环境配置中。

```mermaid
flowchart TD
    Entry[online_test.py 登录与导航] --> Pages[pages 页面入口]
    Pages --> UI[功能 UI 模块]
    Pages --> API[api_client.py HTTP 公共工具]
    UI --> Model[纯数据整理与文件生成模块]
    API --> Main[backend/main.py 认证与接口注册]
    Main --> Routes[documents / esign / contract_history / team_directory]
    Routes --> DB[(MariaDB)]
    Routes --> Storage[服务器文件目录]
```

## 2. Mon équipe：读取、整理、展示分开

| 脚本 | 职责 | 直接依赖 |
| --- | --- | --- |
| `pages/People.py` | 检查登录，按角色选接口，组织页面；不拼 SQL、不生成卡片 HTML | `access_control`、`team_api`、`team_ui`、`organization_ui` |
| `team_api.py` | 发起带认证的读取请求，统一显示接口错误 | `api_client`、Streamlit |
| `team_ui.py` | 将已授权的人员数据转为表格，并展示空列表状态 | `access_control`、Streamlit |
| `organization_model.py` | 纯数据投影：活跃部门、负责人、成员、未分配名单及旧数据冲突提示；无 HTTP、数据库或 Streamlit 依赖 | Python 标准库 |
| `organization_cards.py` | HTML 转义和响应式卡片：宽屏三列、中屏两列、窄屏一列；无部门连线 | Python 标准库 |
| `organization_ui.py` | 默认折叠的架构区域、人数、卡片、未分配名单、冲突提示；无接口请求或虚构数据 | `organization_model`、`organization_cards`、Streamlit |
| `backend/team_directory.py` | 只读 SQL 查询、基础字段投影、两个接口注册；数据库与认证依赖由总入口注入 | FastAPI；不反向导入 `main.py` |

请求流程：

1. 登录后进入 Mon équipe，读取 `/organization-directory`，显示全公司基础架构。架构默认收起。
2. 普通员工读取 `/my-team`；后端从当前登录账号的活跃部门归属查询同组名单，客户端不能指定其他账号或部门扩大范围。
3. Admin、Manager 沿用 `/list-users` 与 `/groups` 的原授权范围，保留现有人员表格。
4. 接口失败时显示错误；不会用不完整的人员数据伪装成全公司架构。

新目录不返回完整邮箱、合同、密码、令牌或财务文件。不增加员工修改权限，也不改变考勤负责人和双签规则。未填写姓名的展示采用邮箱 `@` 前部分；现有目录接口目前只提供该显示名。

## 3. 本地预览与正式功能隔离

`.local-test/organization_preview.py` 是独立预览入口，读取本地只读快照，不接生产 API；账号选择仅模拟名单，不是鉴权测试。虚构布局数据在 `.local-test/organization_fixture.py`，不被正式页面导入，也不部署。该目录、快照和截图被 Git 忽略。

早期树图原型、无引用的 `organization_chart.py` 兼容层及 `streamlit-echarts` 依赖已于 2026-10-16 删除；卡片界面不依赖 ECharts。

## 4. 其他脚本的阅读顺序

| 功能 | 页面入口 | 下一层脚本 |
| --- | --- | --- |
| 考勤 | `pages/Fiches.py` | `fiche_exports.py` 校验生成；`fiche_batch_ui.py` 勾选；`fiche_serialization.py` 保存转换 |
| 日历及月份 | 考勤、Documents 页面 | `calendar_view.py`、`planning.py`、`period_selector.py` |
| 合同 | `pages/Admin.py`、考勤页面 | `contract_ui.py`、`contract_periods.py`、`backend/contract_history.py` |
| Documents | `pages/TransportReceipts.py` | `document_ui.py` 操作；`document_downloads.py` ZIP；`backend/documents.py` 存储与权限 |
| Excel / Word | `fiche_exports.py` | `ExcelGen.py`、`DocxGen.py` 和各自模板 |
| PDF / 双签 | `pdf_export.py`、考勤页面 | `backend/esign.py` 转换发送；`signature_flow.py` 签署人；`signature_placement.py` 签名位置 |
| 人员和组织维护 | `pages/Admin.py` | `backend/main.py` 对应组织接口 |
| 登录和个人资料 | `pages/Login.py`、`pages/Profile.py` | `api_client.py`、`remember_me.py`、后端认证与资料接口 |
| 仪表板及提醒 | `pages/Dashboard.py` | `backend/main.py` 对应提醒与面谈接口 |

`backend/main.py` 仍包含认证、组织管理和提醒，尚未全部拆分。继续拆分时应保持接口路径、权限依赖和旧数据兼容；不需要为一个页面通读整个入口。

## 5. 测试、发布与当前边界

本次新增 `test_team_directory.py`：检查未登录拒绝、三类角色读取、本人部门范围、防止查询参数越权、字段裁剪、数据库连接释放、HTML 转义、冲突数据保留及页面分支。测试使用模拟 API/数据库，无真实上传、邮件或签署。

执行完整回归：`python -m unittest discover -v`。生产发布还需比较服务器差异、备份被替换的源码和 MariaDB、验证平铺模块导入、重启服务并检查健康状态。详细发布记录见 [重构记录](refactoring.md)。

当前是只读展示升级。测试账号的持久化标记/公开隐藏、单部门归属与单负责人强制约束、负责人变更自动调整权限仍待实现；旧冲突数据只提示，不删除。新模块没有数据库结构迁移。

## 2026-10-16 旧版 UI 回退职责映射

下表仅供维护保留的 Streamlit 回退版本。当前 React 导航见 modules.md；不要为新功能从旧页面开始阅读。

| 功能 | 入口 | 职责 |
| --- | --- | --- |
| 导航与登录状态 | `online_test.py` | `access_control.py` 解析前端角色；`remember_me.py` 同步持久登录令牌 |
| 登录及找回密码 | `pages/Login.py` | `api_client.py` 提供 API 地址、认证头、请求异常和错误信息解析 |
| 账户设置 | `pages/Profile.py` | 调用后端资料更新接口 |
| 考勤表创建与编辑 | `pages/Fiches.py` | 组织人员选择、合同绑定、分页、保存与导出操作 |
| 考勤批量选择 | `fiche_batch_ui.py` | 按账号和月份隔离勾选状态，提供全选、取消全选；下载与签署仅处理勾选的表单 |
| 合同编辑界面 | `contract_ui.py` | 管理员维护合同的共用界面 |
| 前端 PDF 转换请求 | `pdf_export.py` | 调用 `/fiches/pdf` 并检查返回内容 |
| Documents 页面 | `pages/TransportReceipts.py` | 上传表单、顶部与文件列表上方联动的月份选择、横向标题和过滤栏、待处理或归档视图、三列卡片 |
| Documents 卡片与操作 | `document_ui.py` | 月份显示、下载准备、归档、删除确认、月份修正、勾选和 ZIP 下载控件 |
| 人员及组织管理 | `pages/Admin.py` | 账户、角色、部门、负责人及合同管理 |
| 团队查看 | `pages/People.py` | Mon équipe：架构默认收起；员工查看本组基础名单，经理与 Admin 的原表格保留 |
| 团队数据请求与名单 | `team_api.py`、`team_ui.py` | 带认证读取与表格整理；普通员工本组名单与管理人员原授权名单分开 |
| 部门架构 | `organization_model.py`、`organization_ui.py` | 纯数据投影与折叠展示分开，无示例数据和预览控制；旧树图适配器已移除 |
| 响应式部门卡片 | `organization_cards.py` | 部门独立卡片与 CSS Grid 自动换行，Direction Générale 位于顶部；部门之间无连线，无拖动、缩放或固定画布高度 |
| 提醒及仪表板 | `pages/Dashboard.py` | 合同到期提醒、年度面谈与设置展示 |
| 调整文件卡片、勾选或显示 | `document_ui.py` |
| 改上传类型或绑定月份 | `pages/TransportReceipts.py`、`backend/documents.py` |
