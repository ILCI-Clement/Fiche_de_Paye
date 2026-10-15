# Presence 新前端

React + TypeScript + Vite + Material UI + TanStack Query。

当前默认运行**真实数据副本的本地重构第一阶段**：`Création de Fiches` 和 `Congés payés`。启动请运行仓库根目录 `scripts/start_local_rebuild.ps1`，访问 `http://127.0.0.1:5173/`，使用原账户密码登录。详情见 [本地重构说明](../doc/local-rebuild.md)。尚未发布。

## 早期独立演示预览（当前未使用）

仓库根目录：

```powershell
python -m uvicorn backend.leave_preview:app --host 127.0.0.1 --port 8012
```

本目录另开终端：

```powershell
npm ci
npm run dev
```

只有显式设置 `VITE_PREVIEW=true` 并将开发代理改为 8012，才会启用此旧演示。当前配置为真实副本和原账户登录，不创建演示员工。

## 入口及维护

- `index.html` → `src/main.tsx` → `src/App.tsx`。
- `src/api.ts`：HTTP、认证头和错误处理；`types.ts`：数据契约。
- `src/components/`：登录、申请、详情/审批及状态标签。
- `theme.ts` / `index.css`：统一主题与电脑/手机布局。
- `npm run build`：类型检查与生产构建；`npm run lint`：静态检查。
- 前端会话存于当前浏览器标签页的 sessionStorage，退出会清理查询缓存；服务器决定真实权限。

## 连接正式后端

覆盖 `.env.local` 中 `VITE_PREVIEW=false` 和 `PRESENCE_API_PROXY`，重启 Vite。仅连接已升级且包含带薪假路由的后端。生产构建不包含演示身份选择；静态部署时反向代理负责将 `/api/` 转发至 FastAPI，并去掉前缀。

API key、密码和数据库配置不得写入任何 `VITE_*` 变量，它们会进入浏览器产物。`leave_preview.py` 只用于 loopback 本地预览，不能作为正式服务。

## Documents preview

PDF/JPG/PNG previews reuse the authenticated file API. `DocumentPreview.tsx` owns
the dialog and in-memory resources; `PdfPreview.tsx` lazily renders PDF pages on
canvas with PDF.js. No third-party preview service or public document URL is used.

Use Node >=22.13 or >=24 for the current PDF.js/Vite toolchain. Vite copies the
installed PDF.js fonts, CMaps and decoders into ignored `public/pdfjs/<version>`
at configuration time for development and production. Deploy the entire `dist`
(including `pdfjs/` and the worker asset), not only JavaScript chunks. Keep older
version directories available while serving already-open clients.
