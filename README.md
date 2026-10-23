# ILCI · Espace RH

当前应用使用 **React + TypeScript 前端与 Python/FastAPI 后端**，涵盖考勤、Documents、请假、员工名册、人员组织、合同及双签。旧 Streamlit 界面已分离为[源码备案](archive/streamlit/README.md)。

- [模块索引与维护说明](doc/modules.md)：按功能定位文件、数据关系、权限及已知限制。
- [脚本职责与应用入口](doc/architecture.md)：当前依赖、运行入口和备案边界。
- [功能概览](doc/overview.md)：业务流程和生产环境说明。
- [后端与部署说明](backend/README.md)：服务启动、文件复制与备份要求。
- [本次整理](doc/source-cleanup.md)：旧版备案、源码精简及检查范围。
- [请假模块设计记录](doc/leave-pilot.md)：React 技术栈、脚本职责、业务规则、本地打开方式与发布准备。
- [本地重构测试版第一阶段](doc/local-rebuild.md)：真实数据副本、独立运行环境、考勤生成、阶段范围及手动测试步骤。

前端入口为 `frontend/src/main.tsx` → `App.tsx` → `AppWorkspace.tsx`；线上 API 入口为 `backend/rebuild_server.py`，共用业务由 `backend/main.py` 注册；独立本地副本入口为 `backend/local_rebuild.py`。启动本地副本使用 `scripts/start_local_rebuild.ps1`。

安装当前 Python 依赖使用 `python -m pip install -r requirements.txt`，开发依赖使用 `requirements-dev.txt`。前端在 `frontend/` 通过 `npm ci` 安装。依赖列表只保留直接使用的软件包，不再默认安装 Streamlit。

历史日志、报告和路径描述可用于查阅过往版本；当前定位以模块索引为准。备案标签为 `archive/streamlit-20261023`，独立源码包及恢复步骤见备案目录。
