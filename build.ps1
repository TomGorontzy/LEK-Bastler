# build.ps1 — Kompatibilitäts-Entrypoint (leitet an src\build.ps1 weiter)

$scriptPath = Join-Path $PSScriptRoot 'src\build.ps1'
if (-not (Test-Path $scriptPath)) {
    Write-Host "Build-Skript nicht gefunden: $scriptPath" -ForegroundColor Red
    exit 1
}

& $scriptPath @args
exit $LASTEXITCODE
