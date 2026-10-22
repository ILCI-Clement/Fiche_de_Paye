$ErrorActionPreference = 'Stop'
$taskRoot = Split-Path -Parent $PSScriptRoot
$taskState = Join-Path $taskRoot '.local-test\rebuild'
if (-not (Test-Path -LiteralPath (Join-Path $taskState 'runtime.json'))) {
    throw 'Local snapshot database must be initialized first.'
}
function Test-LocalPort([int]$Port) {
    # Probe only the requested loopback port; system-wide TCP inventory may be restricted.
    & (Join-Path $taskRoot '.venv\Scripts\python.exe') -c 'import socket,sys; s=socket.socket(); s.settimeout(.5); status=s.connect_ex(("127.0.0.1",int(sys.argv[1]))); s.close(); sys.exit(0 if status==0 else 1)' $Port
    return $LASTEXITCODE -eq 0
}
if (-not (Test-LocalPort 3307)) {
    $localMysqlTemp = Join-Path $taskState 'mysql-tmp'
    New-Item -ItemType Directory -Force -Path $localMysqlTemp | Out-Null
    Start-Process -FilePath 'C:\Program Files\MySQL\MySQL Server 8.4\bin\mysqld.exe' -WindowStyle Hidden -ArgumentList @(
        '--no-defaults', "--datadir=$taskState/mysql-data", '--bind-address=127.0.0.1', '--port=3307',
        '--mysqlx=OFF', "--log-error=$taskState/mysql.log", "--tmpdir=$localMysqlTemp"
    ) | Out-Null
    # Startup recovery may take longer than two seconds; don't start the API early.
    for ($localDbAttempt = 0; $localDbAttempt -lt 30; $localDbAttempt++) {
        if (Test-LocalPort 3307) { break }
        Start-Sleep -Milliseconds 500
    }
    if (-not (Test-LocalPort 3307)) { throw 'Local MySQL did not start. Inspect .local-test/rebuild/mysql.log.' }
}
if (-not (Test-LocalPort 8013)) {
    Start-Process -FilePath (Join-Path $taskRoot '.venv\Scripts\python.exe') -WorkingDirectory $taskRoot -WindowStyle Hidden -ArgumentList @(
        '-m', 'uvicorn', 'backend.local_rebuild:app', '--host', '127.0.0.1', '--port', '8013'
    ) -RedirectStandardOutput (Join-Path $taskState 'api.log') -RedirectStandardError (Join-Path $taskState 'api-errors.log') | Out-Null
}
if (-not (Test-LocalPort 5173)) {
    Start-Process -FilePath (Get-Command node).Source -WorkingDirectory (Join-Path $taskRoot 'frontend') -WindowStyle Hidden -ArgumentList @(
        (Join-Path $taskRoot 'frontend\node_modules\vite\bin\vite.js')
    ) -RedirectStandardOutput (Join-Path $taskState 'frontend.log') -RedirectStandardError (Join-Path $taskState 'frontend-errors.log') | Out-Null
}
Write-Output 'Local test URL: http://127.0.0.1:5173/ — independent DB on port 3307, API on port 8013.'
