# Windows venv python.exe starts a child using the base interpreter. Check the
# workspace launcher as well as the listening child before stopping either app.
function Get-WorkspaceUvicornProcess([int]$Port, [string]$Module, [string]$Workspace) {
    $taskPortErrors = @()
    $taskListeners = @(Get-NetTCPConnection -State Listen -LocalPort $Port -ErrorAction SilentlyContinue -ErrorVariable taskPortErrors)
    if (@($taskPortErrors | Where-Object { $_.CategoryInfo.Category -ne 'ObjectNotFound' }).Count) {
        throw "Cannot inspect port $Port; refusing to report a successful restart."
    }
    $processIds = @($taskListeners | Select-Object -ExpandProperty OwningProcess -Unique)
    if (-not $processIds.Count) { return $null }
    if ($processIds.Count -ne 1) { throw "Multiple processes on port $Port; refusing restart." }
    $apiProcess = Get-CimInstance Win32_Process -Filter "ProcessId=$($processIds[0])"
    $apiCommand = [string]$apiProcess.CommandLine
    $expectedPython = Join-Path $Workspace '.venv\Scripts\python.exe'
    $owned = $apiProcess.ExecutablePath -eq $expectedPython
    if (-not $owned -and $apiProcess.ParentProcessId) {
        $launcher = Get-CimInstance Win32_Process -Filter "ProcessId=$($apiProcess.ParentProcessId)" -ErrorAction SilentlyContinue
        $owned = $launcher -and $launcher.ExecutablePath -eq $expectedPython -and ([string]$launcher.CommandLine).Contains($Module)
    }
    if (-not $owned -or -not $apiCommand.Contains($Module) -or -not $apiCommand.Contains('uvicorn')) {
        throw "Unexpected process on port $Port; refusing restart."
    }
    return $apiProcess
}
