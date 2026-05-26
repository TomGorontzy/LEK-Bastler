# setup.ps1 — Kompatibilitäts-Entrypoint (leitet an src\setup.ps1 weiter)

$scriptPath = Join-Path $PSScriptRoot 'src\setup.ps1'
if (-not (Test-Path $scriptPath)) {
    Write-Host "Setup-Skript nicht gefunden: $scriptPath" -ForegroundColor Red
    exit 1
}

& $scriptPath @args
exit $LASTEXITCODE
