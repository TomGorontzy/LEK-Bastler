# Release Notes v3.6.1

Datum: 2026-05-27

## Highlights

- GUI-Polish: In der Aufgabenliste ist die Spalte `Confidence` jetzt ausgeblendet.
- Stabilität/Typisierung in `src/main.py` verbessert:
  - `tkinter`-`sticky`-Werte typkompatibel vereinheitlicht
  - robuste Typ-Guards für regelbasierte Konfigurationswerte ergänzt
  - sichere Fallbacks für numerische und optionale Rückgabewerte
- Dokumentation aktualisiert und auf `LEK-Bastler` vereinheitlicht:
  - `docs/DOKUMENTATION_TECHNIK.md`
  - `docs/DOKUMENTATION_RELEASES.md`
  - `README.md`

## Qualitätsstatus

- Problems-Check: keine offenen Diagnosen im Workspace
- Regressionstest-Suite: 6/6 grün (`tools/test_regression_core.py`)

## Artefakte

- EXE: `dist/LEK-Bastler_3.6.1/LEK-Bastler_3.6.1.exe`
- Release-Paket: `release/LEK-Bastler_3.6.1.zip`
- Release Notes: `release/RELEASE_NOTES_v3.6.1.md`
