# 可选账户绑定与系统管理员

已发布：`rh-role-handover-20261030T150210Z`，见 [上线验证](releases/role-handover-20261030.md)。

- 登录账号的直接负责人可为空：前后端建号、组织保存均接受无负责人。若明确绑定了失效负责人，仍校验并提示，不自动建立或猜测组织关系。
- 无负责人不产生资料就绪警告；双签仍需有效的实际签署负责人，请假可由 Administration 处理。
- Super Admin 免除员工名册、合同、负责人等员工资料就绪检查。新登记界面选择 Super Admin 后转为仅建系统账号，保留姓名、邮箱和登录安全要求，不生成员工档案或合同。
- Super Admin 编辑界面不再要求绑定员工档案或合同；既有员工、合同、签署和审计历史不删除。升级既有账号不自动清理原有历史或组织关系。
- 授权仍由后端执行，仅 Super Admin 可授予此权限。普通 Admin 无权提升为 Super Admin。

## 调整权限

Super Admin 登录后，在 Administration 打开普通目标账号，使用姓名行右侧 Super Admin 复选框，确认后点击详情窗口右下角 Enregistrer 并完成角色变更确认。

允许 Super Admin 自降级，但必须保留至少一个其他有效 Super Admin；停用或已删除登录的账号不计入。成功后旧会话与恢复凭据全部失效，当前页面清除记住登录并强制重新登录。普通 Admin 仍不得自行撤销 Admin 权限。其他 Super Admin 按既有“仅本人可见”规则隐藏，因此前台没有用于降级其他 Super Admin 的列表入口；此限制不应通过扩大人员目录可见性绕过。

## 验证

- 后端完整回归：564 项通过（含权限移交、自降级、失效后继账号拦截和审计失败回滚），MySQL 仅使用本地临时表，SMTP 模拟。
- 前端 build、lint、测试通过，新增可选负责人、系统账号及保留普通 Responsable 校验的测试。
- 无生产数据修改，无实际权限变更或欢迎邮件发送。

自降级操作：Administration 打开自己 → 取消 Super Admin → 若要普通账户，同时取消 Admin 并选择 Employé / Responsable → 点击详情窗口右下角 Enregistrer 并确认 → 重新登录。仅撤销 Super Admin 时可保留普通 Admin。全部测试使用合成账号；本地和生产 API 均已加载规则。
