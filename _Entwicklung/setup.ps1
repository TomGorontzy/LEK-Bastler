# Setup-Skript für LEK-Bastler

Write-Host "LEK-Bastler Setup" -ForegroundColor Cyan
Write-Host "==================" -ForegroundColor Cyan

# Prüfen ob Python installiert ist
try {
    $pythonVersion = python --version 2>&1
    Write-Host "Python gefunden: $pythonVersion" -ForegroundColor Green
} catch {
    Write-Host "Fehler: Python ist nicht installiert oder nicht im PATH verfügbar." -ForegroundColor Red
    Write-Host "Bitte installieren Sie Python 3.7+ von https://www.python.org/" -ForegroundColor Yellow
    exit 1
}

# Prüfen ob pip verfügbar ist
try {
    $pipVersion = pip --version 2>&1
    Write-Host "pip gefunden: $pipVersion" -ForegroundColor Green
} catch {
    Write-Host "Fehler: pip ist nicht verfügbar." -ForegroundColor Red
    exit 1
}

# Abhängigkeiten installieren
Write-Host "`nInstalliere Abhängigkeiten..." -ForegroundColor Yellow
try {
    pip install -r requirements.txt
    Write-Host "Abhängigkeiten erfolgreich installiert." -ForegroundColor Green
} catch {
    Write-Host "Fehler beim Installieren der Abhängigkeiten." -ForegroundColor Red
    exit 1
}

Write-Host "`nSetup abgeschlossen!" -ForegroundColor Green
Write-Host "Starten Sie die Anwendung mit: python main.py" -ForegroundColor Cyan