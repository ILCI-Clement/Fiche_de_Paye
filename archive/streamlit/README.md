# Streamlit 源码备案

旧界面已从活动开发目录移出。此目录保存可独立解压的源码快照及校验清单；当前应用入口为 React/FastAPI。

## 备案内容

- `source.zip`：完整 Git 源码快照，标签 `archive/streamlit-20261023`，提交 `de055e88d769cd4ad0a01ff51fa05fe07d4e3ef4`。这是已发布签署归档版本之后、员工登记新版之前的基线；旧界面文件与拆分前逐项核对一致。
- `manifest.json`：ZIP 的 SHA-256、大小、提交及移出文件的逐项摘要。
- 包内保留 `online_test.py`、`pages/`、旧日历组件、请求/会话/展示工具、原始依赖清单、模板、配套 API 及旧测试。未保留私有配置、密钥、数据库、业务文件或虚拟环境。
- 这是源码备案。数据库与私有文件应使用各环境的独立备份，密钥另行保管；不能用此 ZIP 覆盖正在写入的生产数据。

## 恢复旧界面进行本地查阅

在仓库根目录使用 PowerShell，解压到一个不存在的新目录，勿覆盖当前工作区：

```powershell
$legacyTarget = Join-Path $PWD '.local-test/streamlit-archive-20261023'
if (Test-Path -LiteralPath $legacyTarget) { throw '请使用新的空目录' }
$legacyManifest = Get-Content archive/streamlit/manifest.json -Raw | ConvertFrom-Json
if ((Get-FileHash archive/streamlit/source.zip -Algorithm SHA256).Hash.ToLowerInvariant() -ne $legacyManifest.sha256) { throw '源码包校验失败' }
Expand-Archive -LiteralPath archive/streamlit/source.zip -DestinationPath $legacyTarget
Set-Location $legacyTarget
python -m venv .venv
.venv/Scripts/python.exe -m pip install -r requirements.txt
New-Item -ItemType Directory -Path .streamlit -Force | Out-Null
'URL_PRESENCE = "http://127.0.0.1:8013"' | Set-Content .streamlit/secrets.toml -Encoding utf8
.venv/Scripts/python.exe -m streamlit run online_test.py --server.address 127.0.0.1 --server.port 8502
```

先在主项目启动独立本地 API（`scripts/start_local_rebuild.ps1`）；以上只运行备案界面，使用本地副本，仍由当前 API 执行权限。恢复目录已由 `.local-test/` 忽略规则保护。

也可通过 Git 标签导出原始源码：`git archive --format=zip --output=streamlit-source.zip archive/streamlit-20261023`。本目录中的 ZIP 本身可用于没有 Git 历史的源码下载包。

不要将备案后端直接作为生产回退入口。线上仍运行的 `presence-interface` 与 flat API 是独立部署，未在此次本地整理中停止；上线时按 `backend/README.md` 比较、备份和分步处理。
