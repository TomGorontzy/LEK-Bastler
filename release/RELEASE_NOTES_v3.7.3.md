# Release Notes v3.7.3
## Enthalten

- LEK-Exportlayout für Aufgabenabstände verfeinert:
  - Zwischen `Überschrift 1` (Aufgabentitel) und Aufgabenbeschreibung wird jetzt konsistent genau eine Leerzeile ausgegeben.
  - Zwischen zwei Aufgaben werden maximal zwei Leerzeilen ausgegeben.
- Export für Aufgaben mit Objekt-/Tabellenstart robuster gemacht:
  - Zusätzliche Leerzeilen vor dem ersten Tabellen-/Objektblock werden vermieden.
  - Führende/abschließende leere Absätze in Aufgabenblöcken werden gezielt begrenzt.

## Qualitätssicherung

- Syntax-/Problems-Check für `src/word_processor.py` ohne Fehler.
- Mini-End-to-End-Exporttest mit tabellenstartender Aufgabe erfolgreich:
  - Titel → eine Leerzeile → Inhalt (inkl. Tabelle/Objekt) bestätigt.

## Artefakte

- dist/LEK-Bastler_3.7.3/LEK-Bastler_3.7.3.exe
- release/LEK-Bastler_3.7.3.zip
- release/RELEASE_NOTES_v3.7.3.md

