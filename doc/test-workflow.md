# 测试与本次精简

## 日常按模块测试

使用项目虚拟环境；只运行本次涉及模块及相关权限、存储测试。例如邮箱改动：

```powershell
.venv/Scripts/python.exe scripts/run_tests.py --mysql test_account_email test_login_guard test_backend_helpers
```

执行完整后端回归，包括已有真实 MySQL 临时表测试：

```powershell
.venv/Scripts/python.exe scripts/run_tests.py --mysql
```

不带 `--mysql` 时遵循已有 `PRESENCE_MYSQL_TEMP_TEST` 环境变量，未启用时会明确显示对应测试 skipped，不能将该次运行当作完整数据库验证。MySQL 测试仍限定本机 3307 的 `presence_local_rebuild` 和连接拥有的 TEMPORARY 表。

脚本默认列出最慢 5 项，可用 `--durations 15` 定位耗时。原来的 `python -m unittest discover -p 'test*.py'` 仍可使用；不引入 pytest 或新依赖。前端测试单独执行 `npm --prefix frontend test`；仅后端改动无需反复构建未改动的前端。

## 2026-10-30 的清理范围

- 将旧 flat API 导入兼容代码备案在 `archive/flat-api/`，25 个活动后端模块统一包导入，去掉 165 行重复兼容分支。
- 移除未被调用的 `fiche_record_belongs_to_user`，以及无操作的 `sync_contract_dates_from_config` 和调用/测试桩。保留真实的账号关联迁移、历史 JSON、合同快照及月份兼容。
- 删除已退役入口的 1 项 flat-import 测试；其他业务、安全及数据库测试全部保留。
- 9 个 SQLite 文件库测试模块共用 `test_support/sqlite_fixtures.py`，连接仅允许内存或操作系统临时目录。一次性文件库使用内存回滚日志并关闭 fsync，仍真实执行约束、提交和回滚；不用于机器崩溃/断电恢复验证。
- 生产数据库连接、日志与持久化配置、真实 MySQL 测试、密码哈希和加密算法均未降低要求。

同一台电脑、相同完整测试配置下的单次对比：清理前 578 项 / 91.638 秒，清理后 577 项 / 47.192 秒，全部通过，约减少 49%。耗时会受系统负载影响，并非固定性能承诺。日志保存在本地 `.local-test/release/test-cleanup-{before,after}.log`。

代码清理已随 `rh-leave-dot-20261030T164037Z` 发布；测试工具与一次性 SQLite 优化仅用于本地开发，不进入生产运行配置。
