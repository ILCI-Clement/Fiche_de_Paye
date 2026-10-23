# 开发导航

先读 `doc/modules.md`，按需求定位模块，再读取相关代码。无需为单个功能通读所有页面或 `backend/main.py`。

## 常用入口

- React 入口：`frontend/src/main.tsx` → `App.tsx` → `AppWorkspace.tsx`；线上 API：`backend/rebuild_server.py`。
- 考勤页面：`frontend/src/attendance/AttendancePage.tsx`；适配：`backend/attendance_api.py`；统一生成：`fiche_exports.py`。
- 新版日历：`frontend/src/attendance/AttendanceCalendar.tsx`、`useCalendarSelection.ts`；共享日期规则仍在 `calendar_view.py`，默认安排在 `planning.py`。
- 合同：`contract_periods.py`、`contract_ui.py`、`backend/contract_history.py`。
- Documents：`frontend/src/documents/`、`backend/documents.py`、`backend/document_numbers.py`；ZIP 共用工具：`document_downloads.py`。
- PDF 与签署：`pdf_export.py`、`backend/esign.py`、`backend/signature_flow.py`、`backend/signature_placement.py`。
- 人员、组织、认证、数据库迁移、提醒：`backend/main.py`；前端对应页面见模块索引。
- React 请求：`frontend/src/api.ts`；后端考勤权限：`backend/attendance_permissions.py`。旧版 Streamlit 已分离至 `archive/streamlit/` 的源码包和 Git 标签；只有需要恢复查阅时才按备案 README 解压，不在活动目录重新引入旧入口。

## 维护约束

- 权限最终由后端执行，页面隐藏按钮不能替代接口鉴权。
- 保留历史合同快照、月份绑定和旧版考勤数据兼容性。重构不得清空数据。
- 保持 Streamlit widget key 稳定，注意 rerun 和 session_state 的修改时机。
- 共用逻辑放在已有对应模块；避免复制校验、文件生成和 API 错误解析。
- 本地与服务器可能有未同步修改；部署前比较差异，按 `backend/README.md` 备份。
- 密钥和密码不得写入源码、文档或命令输出。
- 按用户请求执行测试；静态检查与真实上传、邮件发送、签署要分别说明。
- 改动模块职责、接口或部署文件时，同步更新 `doc/modules.md`。
