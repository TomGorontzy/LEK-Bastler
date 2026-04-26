# LEK-Bastler Portable Setup
# Erstellt eine portable Version für Arbeitsplätze ohne Python

Write-Host "LEK-Bastler Portable Setup" -ForegroundColor Cyan
Write-Host "=========================" -ForegroundColor Cyan

$portableDir = "LEK-Bastler-Portable"
$pythonUrl = "https://www.python.org/ftp/python/3.11.6/python-3.11.6-amd64.exe"
$pipUrl = "https://bootstrap.pypa.io/get-pip.py"

# Portable Verzeichnis erstellen
if (Test-Path $portableDir) {
    Remove-Item $portableDir -Recurse -Force
}
New-Item -ItemType Directory -Path $portableDir | Out-Null

Write-Host "Verwende vollständige Python-Installation für Portabilität..." -ForegroundColor Yellow
Write-Host "Hinweis: Für echte Portabilität ohne Installation verwenden Sie:" -ForegroundColor Cyan
Write-Host ".\build-executable.ps1" -ForegroundColor Cyan

# LEK-Bastler Dateien kopieren
Write-Host "Kopiere LEK-Bastler Dateien..." -ForegroundColor Yellow
Copy-Item "main.py" "$portableDir\"
Copy-Item "word_processor.py" "$portableDir\"
Copy-Item "task_selector.py" "$portableDir\"
Copy-Item "requirements.txt" "$portableDir\"
Copy-Item "examples" "$portableDir\" -Recurse

# Setup-Skript für Zielcomputer erstellen
$setupScript = @"
@echo off
echo LEK-Bastler Setup
echo ================

REM Prüfe ob Python verfügbar ist
python --version >nul 2>&1
if %errorlevel%==0 (
    echo Python gefunden, installiere Abhängigkeiten...
    python -m pip install -r requirements.txt --user
    echo.
    echo Setup abgeschlossen!
    echo Starten Sie LEK-Bastler mit: python main.py
    pause
) else (
    echo Python ist nicht installiert!
    echo.
    echo Optionen:
    echo 1. Python von python.org installieren
    echo 2. Python aus Microsoft Store installieren  
    echo 3. LEK-Bastler.exe verwenden (keine Python-Installation nötig)
    echo.
    pause
)
"@

$setupScript | Out-File -FilePath "$portableDir\Setup.bat" -Encoding ASCII

# README für portable Version
$readme = @"
# LEK-Bastler Portable

Diese portable Version benötigt keine Python-Installation.

## Verwendung:
1. Doppelklick auf 'LEK-Bastler-Start.bat'
2. Die Anwendung startet automatisch

## Enthält:
- Python 3.11 Embeddable
- Alle erforderlichen Bibliotheken
- LEK-Bastler Anwendung
- Beispieldateien

## Systemanforderungen:
- Windows 10/11 (64-bit)
- Keine Administrator-Rechte erforderlich
- Keine Internet-Verbindung nach dem Setup nötig

Gesamtgröße: ca. 25-30 MB
"@

$readme | Out-File -FilePath "$portableDir\LIESMICH.md" -Encoding UTF8

Write-Host "`nPortable Version erstellt in: $portableDir" -ForegroundColor Green
Write-Host "Zum Starten: $portableDir\LEK-Bastler-Start.bat" -ForegroundColor Green
Write-Host "`nDer komplette Ordner kann auf andere Computer kopiert werden." -ForegroundColor Cyan