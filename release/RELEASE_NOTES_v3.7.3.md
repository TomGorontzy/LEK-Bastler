# Release Notes v3.7.3
## Enthalten

- LEK-Exportlayout für Aufgabenabstände verfeinert:
  - Zwischen `Überschrift 1` (Aufgabentitel) und Aufgabenbeschreibung wird jetzt konsistent genau eine Leerzeile ausgegeben.
  - Zwischen zwei Aufgaben werden maximal zwei Leerzeilen ausgegeben.
- Export für Aufgaben mit Objekt-/Tabellenstart robuster gemacht:
  - Zusätzliche Leerzeilen vor dem ersten Tabellen-/Objektblock werden vermieden.
  - Führende/abschließende leere Absätze in Aufgabenblöcken werden gezielt begrenzt.
- Nummerierte Listen im LEK-Export korrigiert:
  - Nummerierte Absätze behalten jetzt pro Aufgabe ihre erwartete Nummerierung und starten nicht mehr versehentlich bei Folgezahlen wie z. B. `4.`.
  - Das Verhalten wurde auch für nummerierte Listen in verschachtelten Tabellen abgesichert.

## Qualitätssicherung

- Syntax-/Problems-Check für `src/word_processor.py` ohne Fehler.
- Mini-End-to-End-Exporttest mit tabellenstartender Aufgabe erfolgreich:
  - Titel → eine Leerzeile → Inhalt (inkl. Tabelle/Objekt) bestätigt.
- Regressionstest-Suite erfolgreich: **27/27 OK** (`tools/test_regression_core.py`).

## Artefakte

- dist/LEK-Bastler_3.7.3/LEK-Bastler_3.7.3.exe
- release/LEK-Bastler_3.7.3.zip
- release/RELEASE_NOTES_v3.7.3.md

