# LEK-Bastler Executable Builder
# Erstellt eine standalone .exe-Datei

Write-Host "LEK-Bastler Executable Builder" -ForegroundColor Cyan
Write-Host "==============================" -ForegroundColor Cyan

# Prüfen ob PyInstaller installiert ist
try {
    $pyinstallerVersion = pip show pyinstaller 2>$null
    if (-not $pyinstallerVersion) {
        Write-Host "Installiere PyInstaller..." -ForegroundColor Yellow
        pip install pyinstaller
    }
} catch {
    Write-Host "Fehler beim Installieren von PyInstaller" -ForegroundColor Red
    exit 1
}

Write-Host "Erstelle standalone executable..." -ForegroundColor Yellow

# Versionsnummer mit Datum und Zeit erstellen
$timestamp = Get-Date -Format "yyyyMMdd_HHmm"
$version = "v2.5_$timestamp"
$exeName = "LEK-Bastler-$version"

# PyInstaller ausführen
pyinstaller --onefile --windowed --name "$exeName" `
    --add-data "examples;examples" `
    --add-data "Vorlagen;Vorlagen" `
    --icon=icon.ico `
    main.py

if ($LASTEXITCODE -eq 0) {
    Write-Host "`nExecutable erfolgreich erstellt!" -ForegroundColor Green
    Write-Host "Datei: dist\$exeName.exe" -ForegroundColor Green
    
    # Kopiere .exe ins LEK-Bastler-Portable Verzeichnis (ohne Versionsnummer)
    if (Test-Path "LEK-Bastler-Portable") {
        Write-Host "Kopiere .exe ins Portable-Verzeichnis..." -ForegroundColor Yellow
        Copy-Item "dist\$exeName.exe" "LEK-Bastler-Portable\LEK-Bastler.exe" -Force
        Write-Host "LEK-Bastler-Portable\LEK-Bastler.exe aktualisiert" -ForegroundColor Green
    }
    
    Write-Host "`nDie .exe-Datei kann direkt auf andere Computer kopiert werden." -ForegroundColor Cyan
    Write-Host "Dateien:" -ForegroundColor Cyan
    Write-Host "  - dist\$exeName.exe (versioniert)" -ForegroundColor Cyan
    Write-Host "  - LEK-Bastler-Portable\LEK-Bastler.exe (aktuell)" -ForegroundColor Cyan
    Write-Host "Größe: ca. 15-20 MB" -ForegroundColor Cyan
    Write-Host "Keine Python-Installation erforderlich!" -ForegroundColor Cyan
} else {
    Write-Host "Fehler beim Erstellen der executable!" -ForegroundColor Red
}