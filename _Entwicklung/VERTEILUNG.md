# Verteilungsoptionen für LEK-Bastler

## 1. Portable Version (Empfohlen)
```powershell
.\create-portable.ps1
```
- **Größe:** ~25-30 MB
- **Vorteile:** Keine Installation, läuft sofort
- **Verwendung:** Ordner kopieren, LEK-Bastler-Start.bat doppelklicken

## 2. Standalone Executable
```powershell
.\build-executable.ps1
```
- **Größe:** ~15-20 MB  
- **Vorteile:** Eine einzige .exe-Datei
- **Verwendung:** LEK-Bastler.exe direkt ausführen

## 3. Netzwerk-Installation
- LEK-Bastler auf Netzlaufwerk bereitstellen
- Benutzer führt portable Version vom Server aus
- Zentrale Updates möglich

## 4. IT-genehmigte Python-Installation
- Python über IT-Abteilung installieren lassen
- Microsoft Store Python (falls erlaubt)
- Conda/Miniconda als Alternative

## Empfehlung nach Umgebung:

### Restriktive Unternehmensumgebung:
→ **Portable Version** (keine Installation nötig)

### Einzelarbeitsplätze:
→ **Standalone .exe** (einfachste Verteilung)

### Bildungseinrichtungen:
→ **Netzwerk-Installation** (zentrale Verwaltung)

### Entwicklungsumgebung:
→ **Standard Python-Installation** (volle Flexibilität)