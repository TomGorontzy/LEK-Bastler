# src\setup.ps1 — Erstellt .venv und installiert Abhängigkeiten aus src\REQUIREMENTS.txt
# Aufruf: .\src\setup.ps1
# Optional: .\src\setup.ps1 -Force  (löscht bestehende .venv und erstellt neu)
param([switch]$Force)

$ErrorActionPreference = 'Stop'

$scriptDir = $PSScriptRoot
$projectDir = (Resolve-Path (Join-Path $scriptDir '..')).Path
$repo = Split-Path $projectDir -Leaf

$venvDir = Join-Path $projectDir '.venv'
$pythonExe = Join-Path $venvDir 'Scripts\python.exe'
$requirementsPath = Join-Path $scriptDir 'REQUIREMENTS.txt'

if ($Force -and (Test-Path $venvDir)) {
    Write-Host "[$repo] Entferne bestehende .venv ..."
    Remove-Item $venvDir -Recurse -Force
}

if (-not (Test-Path $venvDir)) {
    Write-Host "[$repo] Erstelle .venv ..."
    py -m venv $venvDir
} else {
    Write-Host "[$repo] .venv bereits vorhanden."
}

Write-Host "[$repo] Aktualisiere pip ..."
& $pythonExe -m pip install --upgrade pip -q

if (Test-Path $requirementsPath) {
    Write-Host "[$repo] Installiere Pakete aus src\REQUIREMENTS.txt ..."
    & $pythonExe -m pip install -r $requirementsPath -q
    Write-Host "[$repo] Fertig. Aktivieren mit: .\.venv\Scripts\Activate.ps1"
} else {
    Write-Host "[$repo] Keine src\REQUIREMENTS.txt gefunden — .venv angelegt, aber leer."
}
