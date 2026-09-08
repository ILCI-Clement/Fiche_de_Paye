# Start the isolated preview on loopback only.
$ErrorActionPreference = 'Stop'
Set-Location -LiteralPath $PSScriptRoot
$previewPython = Join-Path $PSScriptRoot '.venv\Scripts\python.exe'
if (-not (Test-Path -LiteralPath $previewPython)) {
    throw 'Environnement Python absent. Suivre les instructions du README.'
}
& $previewPython -B -m streamlit run local_preview.py --server.address 127.0.0.1 --server.port 8501 --server.headless true --browser.gatherUsageStats false
