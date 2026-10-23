$ErrorActionPreference = 'Stop'
$localRoot = Split-Path -Parent $PSScriptRoot
. (Join-Path $PSScriptRoot 'local_process.ps1')
$localApi = Get-WorkspaceUvicornProcess 8013 'backend.local_rebuild:app' $localRoot
if ($localApi) {
    Stop-Process -Id $localApi.ProcessId
}
& (Join-Path $PSScriptRoot 'start_local_rebuild.ps1')
