# Release Notes v3.6.1

Datum: 2026-05-27

## Highlights

- GUI-Polish: In der Aufgabenliste ist die Spalte `Confidence` jetzt ausgeblendet.
- Stabilität/Typisierung in `src/main.py` verbessert:
  - `tkinter`-`sticky`-Werte typkompatibel vereinheitlicht
  - robuste Typ-Guards für regelbasierte Konfigurationswerte ergänzt
  - sichere Fallbacks für numerische und optionale Rückgabewerte
- Dokumentation aktualisiert und Markdown-Formatierung bereinigt:
  - `docs/DOKUMENTATION_TECHNIK.md`
  - `memos/MEMO_REGRESSIONSTEST_MATRIX.md`

## Qualitätsstatus

- Problems-Check: keine offenen Diagnosen im Workspace
- Regressionstest-Suite: 6/6 grün (`tests/test_regression_core.py`)

## Artefakte

- Release-Paket: `release/LEK-Bastler-Portable_3.6.1.zip`
