# 2026-10-28 Profile 月份选择界面

- 前端版本：`profile-month-ui-20261028T115926Z`，代码：`3ce60e6eac9fd9f1d843a3107558abaf3308403f`。
- 正式地址：https://app.presence.217.65.145.103.nip.io/ 。
- Nginx 静态目录：`/var/www/presence-react/releases/profile-month-ui-20261028T115926Z/frontend/dist`。
- API 保持 `contact-access-20261028T095545Z` / `f0609e84168b`；`/var/www/presence-react/current` 未切换，后端进程未重启。

## 界面

个人 Profile 使用年份和月份下拉框，只展示所选月份。上半部汇总考勤、文件、签署和请假，下半部展示当月详情。桌面弹窗宽度上限 900px、高度上限 560px，窄屏随窗口适配；切换月份不改变弹窗高度。已归档签署文件沿用共享预览及后端鉴权。

## 验证

- 52 项前端测试、TypeScript、lint 和生产构建通过。
- 本地合成 Responsable 数据核对月份切换与单月内容，390px 及桌面未出现横向溢出。
- 250 个新静态文件、8 个线上资源哈希匹配；300 个上一版文件保持不变。
- 10 个有效账号的本人账号/Profile 共 20 次只读接口检查通过；3 个匿名私有接口请求被拒绝。
- 使用线上已登录 Test Responsable 打开本人 Profile，核对十月、九月、十月切换；仅一个月度面板，900×560px 弹窗无横向溢出。
- 四项服务保持 active，API 进程未变化；未修改业务数据、发送真实邮件或签署请求。

## 备份与回退

受保护备份：`/var/backups/presence-react/profile-month-ui-20261028T115926Z`，包含前端、Nginx 配置、数据库、源码/配置及 keyring、私有文件和签署 SQLite。密钥未写入仓库。旧哈希静态资源保留以兼容已打开页面。

仅回退本次界面，保持 API 与业务数据：

```sh
sudo /var/backups/presence-react/profile-month-ui-20261028T115926Z/rollback-ui.sh
```

API 的协调回退另见 [上一版发布记录](account-contact-20261028.md)，不要将本次仅前端目录用作 API 发布目录。
