$ErrorActionPreference = 'Stop'
$gatewayRoot = Split-Path -Parent $PSScriptRoot
$gatewayState = Join-Path $gatewayRoot '.local-test/rebuild'
$savedTunnelFile = Join-Path $gatewayState 'pdf-tunnel.json'
if (-not (Test-Path -LiteralPath $savedTunnelFile)) { Write-Output 'No recorded PDF tunnel.'; exit 0 }
$savedTunnel = Get-Content -LiteralPath $savedTunnelFile -Raw | ConvertFrom-Json
$cloudProcess = Get-CimInstance Win32_Process -Filter "ProcessId=$($savedTunnel.tunnel_pid)" -ErrorAction SilentlyContinue
if ($cloudProcess -and $cloudProcess.ExecutablePath -eq (Join-Path $gatewayState 'cloudflared.exe') -and $cloudProcess.CommandLine -like '*127.0.0.1:8014*') { Stop-Process -Id $cloudProcess.ProcessId }
$gatewayProcess = Get-CimInstance Win32_Process -Filter "ProcessId=$($savedTunnel.gateway_pid)" -ErrorAction SilentlyContinue
if ($gatewayProcess -and $gatewayProcess.CommandLine -like '*uvicorn*backend.pdf_test_gateway:app*') { Stop-Process -Id $gatewayProcess.ProcessId }
$runtimeFile = Join-Path $gatewayState 'runtime.json'
$gatewayRuntime = Get-Content -LiteralPath $runtimeFile -Raw | ConvertFrom-Json
$gatewayRuntime.PSObject.Properties.Remove('esign_public_url')
$gatewayRuntime | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $runtimeFile -Encoding utf8
Remove-Item -LiteralPath $savedTunnelFile
Write-Output 'PDF-only tunnel stopped. Restart the local backend to refresh signing readiness.'
