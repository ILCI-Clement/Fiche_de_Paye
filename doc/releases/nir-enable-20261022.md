# 2026-10-22 生产社保号录入配置

- 配置版本：`nir-enable-20261022T140534Z`，用户明确要求开放线上 Admin / SuperAdmin 录入与查看。
- 应用源码未改，仍为 `7ed872e`；此前本地代码清理未夹带部署。
- 只新增 React API 的 systemd drop-in `40-personnel-nir.conf`，设置 `PRESENCE_PERSONNEL_NIR_ENABLED=1` 和独立 keyring 路径，重启 `presence-react-api`。
- 新生产密钥由密码学随机源生成，不复用本地或会话密钥；未录入真实号码，未修改员工、密码或角色。

## 加密与访问

后端将规范化号码以 AES-256-GCM 加密后存储。每次生成独立 12 字节随机 nonce；认证附加数据绑定员工 ID、用途、格式版本及 key ID。数据库只保存版本、算法、key ID、nonce 和含认证标签的密文，不能跨员工复制后解密。

keyring 位于 `/var/lib/presence-secrets/personnel-nir.json`，目录为 root:streamlitapp 0750，文件 0640；Web 静态目录及数据库备份不包含该文件。密钥单独备份在 root 专用 `/var/backups/presence-nir-keys/nir-enable-20261022T140534Z.json`，目录 0700、文件 0600。该备份仍位于同一服务器，运维应按密钥备份策略另行保存受保护的异机副本；丢失密钥将无法解密历史号码。轮换必须保留仍被密文或备份引用的旧 key。

仅有效 Admin / SuperAdmin 可录入或专用查看。查看先提交审计，响应 no-store；明文不进入普通 Profile、查询缓存或浏览器存储，前端 30 秒或离开页面即隐藏。输入框遮挡显示本身不承担加密。此加密只覆盖社保号字段，不覆盖普通姓名、地址、自由文本或附件。

## 配置核对与关闭

- 公网能力接口：2 个现有管理员账号启用成功；5 个非管理员返回 403，匿名返回 401。
- 已检查 keyring 有效性、服务账号读取权限、独立密钥备份一致性及 API 启动；未运行新的测试套件、写入生产测试号码或重新执行浏览器录入验收。
- 原数据库及配置备份：`/var/backups/presence-react/nir-enable-20261022T140534Z`。

关闭录入和解密查看：

```sh
sudo /var/backups/presence-react/nir-enable-20261022T140534Z/disable-nir.sh
```

关闭只改变能力开关并重启 React API；保留密文、keyring 和密钥备份，不能为回退而删除可能已被新记录引用的密钥。
