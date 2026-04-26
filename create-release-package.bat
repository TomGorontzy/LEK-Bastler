@echo off
REM LEK-Bastler Portable - Release Package Creator
REM Erstellt ein ZIP-Archiv mit allen benötigten Dateien für die Distribution

echo ================================================
echo LEK-Bastler Portable - Release Package Creator
echo ================================================
echo.

REM Arbeitsverzeichnis setzen
cd /d "%~dp0"

REM Prüfen ob LEK-Bastler-Portable Verzeichnis existiert
if not exist "LEK-Bastler-Portable" (
    echo FEHLER: LEK-Bastler-Portable Verzeichnis nicht gefunden!
    echo Bitte zuerst das portable Build erstellen.
    pause
    exit /b 1
)

REM Temporäres Release-Verzeichnis erstellen
set TEMP_DIR=LEK-Bastler-Release-Temp
set RELEASE_NAME=LEK-Bastler

echo Erstelle temporäres Release-Verzeichnis...
if exist "%TEMP_DIR%" rmdir /s /q "%TEMP_DIR%"
mkdir "%TEMP_DIR%"
mkdir "%TEMP_DIR%\%RELEASE_NAME%"

echo.
echo Kopiere benötigte Dateien...

REM Hauptprogramm
echo - LEK-Bastler.exe
copy "LEK-Bastler-Portable\LEK-Bastler.exe" "%TEMP_DIR%\%RELEASE_NAME%\"

REM Dokumentation
echo - LIESMICH.md
copy "LEK-Bastler-Portable\LIESMICH.md" "%TEMP_DIR%\%RELEASE_NAME%\"

REM Komplette Dokumentation
echo - Dokumentation\
mkdir "%TEMP_DIR%\%RELEASE_NAME%\Dokumentation"
xcopy "LEK-Bastler-Portable\Dokumentation\*" "%TEMP_DIR%\%RELEASE_NAME%\Dokumentation\" /e /y

# Vorlagen
echo - Vorlagen\
mkdir "%TEMP_DIR%\%RELEASE_NAME%\Vorlagen"
xcopy "LEK-Bastler-Portable\Vorlagen\*" "%TEMP_DIR%\%RELEASE_NAME%\Vorlagen\" /e /y

REM Aufgaben (Beispieldateien)
echo - Aufgaben\
mkdir "%TEMP_DIR%\%RELEASE_NAME%\Aufgaben"
xcopy "LEK-Bastler-Portable\Aufgaben\*" "%TEMP_DIR%\%RELEASE_NAME%\Aufgaben\" /e /y

REM LEKs Verzeichnis
echo - LEKs\
mkdir "%TEMP_DIR%\%RELEASE_NAME%\LEKs"
copy "LEK-Bastler-Portable\LEKs\LIESMICH.md" "%TEMP_DIR%\%RELEASE_NAME%\LEKs\"

echo.
echo Erstelle ZIP-Archiv...

REM PowerShell für ZIP-Erstellung verwenden
powershell -Command "Compress-Archive -Path '%TEMP_DIR%\%RELEASE_NAME%' -DestinationPath '%RELEASE_NAME%.zip' -Force"

if errorlevel 1 (
    echo FEHLER: ZIP-Erstellung fehlgeschlagen!
    echo Fallback: Versuche mit 7-Zip...
    
    REM Fallback: 7-Zip verwenden falls PowerShell fehlschlägt
    if exist "C:\Program Files\7-Zip\7z.exe" (
        "C:\Program Files\7-Zip\7z.exe" a -tzip "%RELEASE_NAME%.zip" "%TEMP_DIR%\%RELEASE_NAME%\*"
    ) else if exist "C:\Program Files (x86)\7-Zip\7z.exe" (
        "C:\Program Files (x86)\7-Zip\7z.exe" a -tzip "%RELEASE_NAME%.zip" "%TEMP_DIR%\%RELEASE_NAME%\*"
    ) else (
        echo FEHLER: Weder PowerShell noch 7-Zip verfügbar!
        echo Manuell komprimieren: %TEMP_DIR%\%RELEASE_NAME%
        pause
        exit /b 1
    )
)

REM Aufräumen
echo Räume temporäre Dateien auf...
rmdir /s /q "%TEMP_DIR%"

echo.
echo ================================================
echo Release Package erfolgreich erstellt!
echo ================================================
echo Datei: %RELEASE_NAME%.zip
echo Größe:
dir "%RELEASE_NAME%.zip" | findstr "%RELEASE_NAME%.zip"
echo.
echo Inhalt:
echo - LEK-Bastler.exe (Hauptprogramm)
echo - LIESMICH.md (Anleitung)
echo - Dokumentation\ (Technische Dokumentation)
echo - Vorlagen\ (LEK-Vorlagen)
echo - Aufgaben\ (Beispiel-Aufgabendateien)
echo - LEKs\ (Export-Verzeichnis)
echo.
echo Das ZIP-Archiv ist bereit für die Verteilung!
echo ================================================

pause