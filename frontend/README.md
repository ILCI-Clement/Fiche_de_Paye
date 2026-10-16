# Presence React 前端

React + TypeScript + Vite + Material UI + TanStack Query，已于 2026-10-15 上线。
入口：`index.html` → `src/main.tsx` → `src/App.tsx` → `src/AppWorkspace.tsx`。
AppWorkspace 统一导航，各业务页只负责本模块，不再保留独立请假应用外壳。

## 本地运行

在仓库根目录运行 `scripts/start_local_rebuild.ps1`，打开 http://127.0.0.1:5173/。
使用已有真实账户副本登录，连接独立本地后端 8013 与数据库副本，详见
[本地重构说明](../doc/local-rebuild.md)。不再提供虚构账号选择或 8012 演示后端。

- `npm ci` 安装依赖；`npm run dev` 开发；`npm run build` 类型检查与构建；`npm run lint` 静态检查。
- 本地代理由 `.env.local` 中 `PRESENCE_API_PROXY` 设置，默认独立本地后端。
- 生产部署完整 `dist/`，同源反向代理将 `/api/` 转发至正式 API 并去掉前缀。
- API key、密码及数据库配置不得写入 `VITE_*` 浏览器变量。
- 登录/认证集中在 `api.ts`、`authSession.ts`；主题在 `theme.ts`、`index.css`。
- 各业务入口与草稿保护见 [模块索引](../doc/modules.md)。

## Documents preview

PDF/JPG/PNG previews reuse the authenticated file API. `DocumentPreview.tsx` owns
the dialog and in-memory resources; `PdfPreview.tsx` lazily renders PDF pages on
canvas with PDF.js. No third-party preview service or public document URL is used.

Use Node >=22.13 or >=24 for the current PDF.js/Vite toolchain. Vite copies the
installed PDF.js fonts, CMaps and decoders into ignored `public/pdfjs/<version>`
at configuration time for development and production. Deploy the entire `dist`
(including `pdfjs/` and the worker asset), not only JavaScript chunks. Keep older
version directories available while serving already-open clients.
