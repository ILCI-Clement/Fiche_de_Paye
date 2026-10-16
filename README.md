# Presence 项目导航

当前线上使用 **React + TypeScript 前端与 Python/FastAPI 后端**，涵盖考勤、Documents、请假、人员组织、合同及双签。Streamlit 仅作为保留的回退版本，不再是新功能入口。

- [模块索引与维护说明](doc/modules.md)：按功能定位文件、数据关系、权限及已知限制。
- [脚本职责与应用入口](doc/architecture.md)：运行入口、模块依赖、旧版回退边界及发布方式。
- [Documents 开发报告](doc/documents-report.md)：新文件模块的功能、操作、存储权限、上线状态与后续优化。
- [功能概览](doc/overview.md)：业务流程和生产环境说明。
- [后端与部署说明](backend/README.md)：服务启动、文件复制与备份要求。
- [重构记录](doc/refactoring.md)：本次代码整理范围及检查结果。
- [代码优化与维护计划](doc/code-audit.md)：代码依据、风险优先级、瘦身顺序及后续验证入口。
- [历史 Mon équipe 架构预览](doc/organization-preview.md)：本地效果、数据来源和待实现的权限规则。
- [请假模块设计记录](doc/leave-pilot.md)：React 技术栈、脚本职责、业务规则、本地打开方式与发布准备。
- [本地重构测试版第一阶段](doc/local-rebuild.md)：真实数据副本、独立运行环境、考勤生成、阶段范围及手动测试步骤。
- [历史开发日志](doc/dev-log.md)：过往改动。

前端入口为 `frontend/src/main.tsx` → `App.tsx` → `AppWorkspace.tsx`；线上 API 入口为 `backend/rebuild_server.py`，共用业务由 `backend/main.py` 注册；独立本地副本入口为 `backend/local_rebuild.py`。启动本地副本使用 `scripts/start_local_rebuild.ps1`。

2026-10-16 已删除无人引用的树图原型、虚构账号演示与重复请假外壳，并停止跟踪编译缓存。旧版页面 `online_test.py` / `pages/` 及共享业务模块仍用于回退，保留原因见 `doc/architecture.md`。清理前 Git 标签：`checkpoint/before-legacy-cleanup-20261016`。
