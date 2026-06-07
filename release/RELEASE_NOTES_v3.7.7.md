# Release Notes v3.7.7

## Enthalten

- GUI-Layout im erweiterten Modus verbessert:
  - Die Fensterhöhe wird jetzt bei Bedarf automatisch an den sichtbaren Inhalt angepasst.
  - Dadurch bleiben die unteren Schaltflächen auch dann sichtbar, wenn zusätzliche erweiterte Funktionen eingeblendet werden.
  - Der Mechanismus greift sowohl beim Programmstart als auch beim Umschalten des Modus.

## Qualitätssicherung

- Syntax-/Problems-Check für `src/main.py` erfolgreich: keine Fehler gefunden.
- Regressionstest-Suite erfolgreich: **33/33 OK** (`tools/test_regression_core.py`).

## Artefakte

- dist/LEK-Bastler_3.7.7/LEK-Bastler_3.7.7.exe
- release/LEK-Bastler_3.7.7.zip
- release/RELEASE_NOTES_v3.7.7.md
