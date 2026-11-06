# 员工名册历史快照（本地版本）

入口：Registre du personnel 顶部 **Historique des archives**。Admin/Super Admin 可查看、手动存档、重试和下载；仅 Super Admin 可修改周期。没有删除接口，不自动清理旧档。

## 已确认的范围

- 只含采集时未归档的员工及其全部任职经历，每段一行，沿用现有 Excel 模板。外部登录账号没有名册档案，不进入快照。零人时保存仅表头的快照。
- 默认半年：Europe/Paris 每年 1/7 月首日 03:00；可选 1/3/6/12 个月，按自然周期对齐。数据库时间为 UTC，界面转法国时间。首次启用只设置未来计划，不生成、不追溯。
- 手动存档不改变计划。停机错过多个周期只补一份恢复时状态，分别记录原计划和实际采集时间；不能还原停机期间的历史状态。修改周期不抹掉已到期的任务。
- worker 每 60 秒检查；自动失败每小时重试，崩溃留下的自动 running 任务可恢复。GET_LOCK 跨进程互斥，自动计划时间唯一；成功档案不重新生成。手动请求 UUID 幂等。
- 失败进入首页 Admin 的“Archives à vérifier”，可打开快照窗口重试。

## 保存、下载与审计

使用 REPEATABLE READ / WITH CONSISTENT SNAPSHOT 一次读取人员与任职资料，生成“含社保号”和“不含社保号”两个 Excel，分别 AES-256-GCM 加密后存入 MySQL LONGBLOB，无明文临时文件。下载只解密已有版本，不再读取现有员工资料。文件内元数据及文件名含实际采集时间。

AAD 绑定用途、快照 UUID、版本种类和 key id，另校验 Excel SHA-256。生成内容与审计原子提交，失败不发布部分档案。沿用 10,000 行上限，超限失败，不静默截断。两份工作簿都整体加密。

下载默认不含社保号；勾选完整号码时明确提示下载后为明文 Excel，后端重新验证当前 Admin 权限和确认字段。准备、生成、失败、重试、下载及周期修改写入独立审计，记录身份/时间/是否含号码，不记录号码、密文或密钥。接口不缓存，也不返回密文。

## 生产配置：尚未发布

- **必须配置独立生产密钥**：`PRESENCE_PERSONNEL_SNAPSHOT_KEYS_FILE`，key-ring 格式为 active_key_id 与 keys（32 字节随机 AES 密钥的 base64url）。格式解析共用 NIR 逻辑，但文件、密钥及用途分离；不自动生成生产密钥。
- `PRESENCE_PERSONNEL_SNAPSHOT_WORKER=1`（默认）启动 worker。其他进程可设置 0，多个进程开启也受数据库互斥保护。
- 备份需包含 `personnel_snapshot_settings`、`personnel_snapshots`（含密文 BLOB）、`personnel_snapshot_events`，建议 mysqldump 使用 `--hex-blob`。密钥文件须受保护地另行备份并验证恢复，不放源码、公开目录或普通 SQL 备份内。
- 轮换只改变 active key，保留全部历史 key；丢失旧密钥会使相应旧档不可解密。恢复须同时核对密文、审计和历史密钥备份。
- 本地独立密钥位于 Git 忽略的 `.local-test/rebuild/keys/personnel-snapshots.json`；不复用生产/NIR key。首次下一次为 2027-01-01 03:00 法国时间；未生成真实员工快照。

## 模块与测试

`backend/personnel_snapshots.py`：增量表、存档、计划、Admin 路由及审计；`snapshot_schedule.py`：日期/DST；`snapshot_crypto.py`：独立 key ring 与文件加密。`frontend/src/personnel/PersonnelSnapshots.tsx` / `snapshotApi.ts`：独立窗口、20 条分页、下载确认与周期冲突检查。

接口：GET/POST `/personnel-snapshots`；PUT `/personnel-snapshots/settings`；POST `/personnel-snapshots/{uuid}/retry`、`/download`。只注册到新版 API 和本地入口。

测试用 MySQL 连接级 TEMPORARY 表和合成号码，覆盖 DST、并发锁、重复任务、空表、多任职、归档排除、历史不变、下载选择、权限撤销、周期冲突、密钥轮换、篡改、超限、失败重试和补偿。浏览器验收只打开空窗口，未创建真实资料快照或发送邮件。
