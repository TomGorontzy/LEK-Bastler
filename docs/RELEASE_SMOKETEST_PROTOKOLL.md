# Release-Smoketest-Protokoll

## Zweck

Nachvollziehbare, wiederholbare Dokumentation der Release-Smoketests je Version.

## Prüfschema (pro Release)

1. Integrität des Release-ZIP-Artefakts (Existenz, Größe, SHA256)
2. Entpacktest in isoliertes Temp-Verzeichnis
3. Pflichtinhalte vorhanden
4. EXE-Startprobe bis GUI-Idle
5. Ergebnisse + Auffälligkeiten dokumentieren

## Protokoll: v3.6.1 (2026-05-27)

### Eingangsdaten

- Artefakt: `release/LEK-Bastler_3.6.1.zip`
- Größe: `15.73 MB`
- SHA256: `2A3BC0F88FBA2AFBE8190EABADDA39F297113767D882E63F1AB0E1C598026E8E`

### Entpack- und Inhaltsprüfung

- Entpackziel: `%TEMP%/LEK-Bastler-smoke-3.6.1`
- Archivstruktur: Dateien/Ordner direkt im ZIP-Root (kein zusätzlicher Top-Level-Unterordner)
- Pflichtinhalte geprüft: OK
  - `LEK-Bastler_3.6.1.exe`
  - `README.md`
  - `LIZENZ.txt`
  - `data/`
  - `docs/`
  - `data/LEKs/README.md`

### Startprobe EXE

- Prozessstart: OK (`INPUT_IDLE=True`)
- Test-PID: `12204`
- Prozess danach kontrolliert beendet: OK

### Ergebnis

- **Smoketest bestanden**
- Keine technischen Auffälligkeiten im Post-Release-Check

## Protokoll: v3.6.2 (2026-05-31)

### Eingangsdaten

- Artefakt: `release/LEK-Bastler_3.6.2.zip`
- Größe: `15.75 MB`
- SHA256: `D214DEDE23E399DD5899E1BCA5DC211241AE0DFB602ABF28ADC0A3814726B0FD`

### Entpack- und Inhaltsprüfung

- Entpackziel: `%TEMP%/LEK-Bastler-smoke-3.6.2`
- Archivstruktur: Dateien/Ordner direkt im ZIP-Root (kein zusätzlicher Top-Level-Unterordner)
- Pflichtinhalte geprüft: OK
  - `LEK-Bastler_3.6.2.exe`
  - `README.md`
  - `LIZENZ.txt`
  - `data/`
  - `docs/`
  - `data/LEKs/README.md`

### Startprobe EXE

- Prozessstart: OK (`INPUT_IDLE=True`)
- Test-PID: `7688`
- Prozess danach kontrolliert beendet: OK

### Ergebnis

- **Smoketest bestanden**
- Keine technischen Auffälligkeiten im Post-Release-Check

## Protokoll: v3.7.0 (2026-06-06)

### Eingangsdaten

- Artefakt: `release/LEK-Bastler_3.7.0.zip`
- Größe: `15.71 MB` (`16,475,911` Bytes)
- SHA256: `A8BAB86D1320F3F9DE255CB8B19DCC06842CE03635D64D2912CD9EB532D83CAD`

### Entpack- und Inhaltsprüfung

- Entpackziel: `%TEMP%/LEK-Bastler-smoke-3.7.0-20260606_163101`
- Archivstruktur: Dateien/Ordner direkt im ZIP-Root (kein zusätzlicher Top-Level-Unterordner)
- Pflichtinhalte geprüft: OK
  - `LEK-Bastler_3.7.0.exe`
  - `README.md`
  - `LIZENZ.txt`
  - `data/`
  - `docs/`
  - `data/LEKs/README.md`

### Startprobe EXE

- Prozessstart: OK (`INPUT_IDLE=True`)
- Test-PID: `16576`
- Prozess danach kontrolliert beendet: OK

### Ergebnis

- **Smoketest bestanden**
- Keine technischen Auffälligkeiten im Post-Release-Check

### Post-Release (GitHub)

- Release veröffentlicht/aktualisiert unter: <https://github.com/TomGorontzy/LEK-Bastler/releases/tag/v3.7.0>
- Enthaltenes Asset: `LEK-Bastler_3.7.0.zip`
