# Presence 项目导航

现有 Streamlit 前端与 FastAPI 后端，用于人员管理、月度考勤、文件凭证和文档导出。`frontend/` 新增 React 带薪假申请试点，先在本地独立体验。

- [模块索引与维护说明](doc/modules.md)：按功能定位文件、数据关系、权限及已知限制。
- [脚本职责与应用入口](doc/architecture.md)：两个总入口、模块依赖、Mon équipe 数据流及测试发布方式。
- [Documents 开发报告](doc/documents-report.md)：新文件模块的功能、操作、存储权限、上线状态与后续优化。
- [功能概览](doc/overview.md)：业务流程和生产环境说明。
- [后端与部署说明](backend/README.md)：服务启动、文件复制与备份要求。
- [重构记录](doc/refactoring.md)：本次代码整理范围及检查结果。
- [代码优化与维护计划](doc/code-audit.md)：代码依据、风险优先级、瘦身顺序及后续验证入口。
- [Mon équipe 架构预览](doc/organization-preview.md)：本地效果、数据来源和待实现的权限规则。
- [新版带薪假申请试点](doc/leave-pilot.md)：React 技术栈、脚本职责、业务规则、本地打开方式与发布准备。
- [本地重构测试版第一阶段](doc/local-rebuild.md)：真实数据副本、独立运行环境、考勤生成、阶段范围及手动测试步骤。
- [历史开发日志](doc/dev-log.md)：过往改动。

原前端入口为 `online_test.py`，新前端入口为 `frontend/src/main.tsx`，正式后端入口为 `backend/main.py`；独立试点后端为 `backend/leave_preview.py`。先根据模块索引定位代码，再阅读相应文件。
