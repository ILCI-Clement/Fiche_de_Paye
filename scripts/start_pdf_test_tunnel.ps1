$ErrorActionPreference = 'Stop'
$gatewayRoot = Split-Path -Parent $PSScriptRoot
. (Join-Path $PSScriptRoot 'local_process.ps1')
$gatewayState = Join-Path $gatewayRoot '.local-test/rebuild'
$gatewayPython = Join-Path $gatewayRoot '.venv/Scripts/python.exe'
$cloudBinary = Join-Path $gatewayState 'cloudflared.exe'
if (-not (Test-Path -LiteralPath $cloudBinary)) { throw 'Download and verify official cloudflared first.' }
$gatewayProcess = Get-WorkspaceUvicornProcess 8014 'backend.pdf_test_gateway:app' $gatewayRoot
if ($gatewayProcess) {
    Stop-Process -Id $gatewayProcess.ProcessId
}
# Reload the gateway too: an existing PDF-only process cannot accept callbacks.
$gatewayProcess = Start-Process -FilePath $gatewayPython -WorkingDirectory $gatewayRoot -WindowStyle Hidden -PassThru -ArgumentList @(
        '-m', 'uvicorn', 'backend.pdf_test_gateway:app', '--host', '127.0.0.1', '--port', '8014', '--no-access-log'
    ) -RedirectStandardOutput (Join-Path $gatewayState 'pdf-gateway.log') -RedirectStandardError (Join-Path $gatewayState 'pdf-gateway-errors.log')
$gatewayPid = $gatewayProcess.Id
function Set-SigningGatewayUrl([string]$SigningPublicUrl) {
    $runtimeFile = Join-Path $gatewayState 'runtime.json'
    $gatewayRuntime = Get-Content -LiteralPath $runtimeFile -Raw | ConvertFrom-Json
    $gatewayRuntime | Add-Member -NotePropertyName esign_public_url -NotePropertyValue $SigningPublicUrl -Force
    $gatewayRuntime | Add-Member -NotePropertyName esign_callback_url -NotePropertyValue "$SigningPublicUrl/esign/callback" -Force
    $gatewayRuntime | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $runtimeFile -Encoding utf8
    & (Join-Path $PSScriptRoot 'restart_local_rebuild.ps1')
}
$savedTunnelFile = Join-Path $gatewayState 'pdf-tunnel.json'
if (Test-Path -LiteralPath $savedTunnelFile) {
    $savedTunnel = Get-Content -LiteralPath $savedTunnelFile -Raw | ConvertFrom-Json
    $savedProcess = Get-CimInstance Win32_Process -Filter "ProcessId=$($savedTunnel.tunnel_pid)" -ErrorAction SilentlyContinue
    if ($savedProcess -and $savedProcess.ExecutablePath -eq $cloudBinary -and $savedProcess.CommandLine -like '*127.0.0.1:8014*') {
        $savedTunnel.gateway_pid = $gatewayPid
        $savedTunnel | ConvertTo-Json | Set-Content -LiteralPath $savedTunnelFile -Encoding utf8
        Set-SigningGatewayUrl $savedTunnel.public_url
        Write-Output "Existing signing test tunnel (PDF + callback): $($savedTunnel.public_url)"
        exit 0
    }
}
$tunnelErrors = Join-Path $gatewayState 'pdf-tunnel-errors.log'
$tunnelProcess = Start-Process -FilePath $cloudBinary -WorkingDirectory $gatewayState -WindowStyle Hidden -PassThru -ArgumentList @(
    'tunnel', '--no-autoupdate', '--url', 'http://127.0.0.1:8014', '--protocol', 'http2'
) -RedirectStandardOutput (Join-Path $gatewayState 'pdf-tunnel.log') -RedirectStandardError $tunnelErrors
$publicUrl = $null
for ($tunnelIndex = 0; $tunnelIndex -lt 30; $tunnelIndex++) {
    Start-Sleep -Seconds 1
    if (Test-Path -LiteralPath $tunnelErrors) {
        $urlMatch = [regex]::Match((Get-Content -LiteralPath $tunnelErrors -Raw), 'https://[a-z0-9-]+\.trycloudflare\.com')
        if ($urlMatch.Success) { $publicUrl = $urlMatch.Value; break }
    }
    if ($tunnelProcess.HasExited) { throw 'Temporary tunnel stopped; inspect its private log.' }
}
if (-not $publicUrl) { throw 'Temporary tunnel URL not available; inspect its private log.' }
@{ tunnel_pid = $tunnelProcess.Id; gateway_pid = $gatewayPid; public_url = $publicUrl } | ConvertTo-Json | Set-Content -LiteralPath $savedTunnelFile -Encoding utf8
Set-SigningGatewayUrl $publicUrl
Write-Output "Signing test URL (PDF + callback): $publicUrl"
Write-Output 'Local backend restarted. Keep the tunnel running until final PDF archiving completes. Stop with scripts/stop_pdf_test_tunnel.ps1.'
