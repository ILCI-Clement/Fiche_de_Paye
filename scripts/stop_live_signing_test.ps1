$ErrorActionPreference = 'Stop'
$signingRoot = Split-Path -Parent $PSScriptRoot
$signingRuntimePath = Join-Path $signingRoot '.local-test/rebuild/runtime.json'
$signingRuntime = Get-Content -LiteralPath $signingRuntimePath -Raw | ConvertFrom-Json
if ($signingRuntime.PSObject.Properties.Name -contains 'live_signing_test') { $signingRuntime.live_signing_test.enabled = $false }
$signingRuntime.PSObject.Properties.Remove('live_test_key')
$signingRuntime | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $signingRuntimePath -Encoding utf8
$signingListeners = Get-NetTCPConnection -State Listen -LocalPort 8013 -ErrorAction SilentlyContinue
foreach ($signingListener in $signingListeners) {
    $signingProcess = Get-CimInstance Win32_Process -Filter "ProcessId=$($signingListener.OwningProcess)"
    if ($signingProcess.CommandLine -notlike '*uvicorn*backend.local_rebuild:app*') { throw 'Unexpected process on local backend port.' }
    Stop-Process -Id $signingProcess.ProcessId
}
& (Join-Path $PSScriptRoot 'start_local_rebuild.ps1')
Write-Output 'Live signing test disabled; sandbox restored.'
