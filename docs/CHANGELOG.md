# Changelog

Alle relevanten Änderungen an diesem Projekt werden hier dokumentiert.

## [Unreleased]

## [3.6.2] - 2026-05-31

### Changed

- Dokumentation und Release-Begriffe redaktionell vereinheitlicht.
- GitHub-Issue-Status auf abgeschlossene Sprints und anstehende Anwenderprüfung ausgerichtet.

### Added

- Vollständige Anwenderprüfungs-Checkliste unter `docs/ANWENDERPRUEFUNG_CHECKLISTE.md` ergänzt.
- Kompakte 5–10-Minuten-Kurzcheckliste unter `docs/ANWENDERPRUEFUNG_KURZCHECKLISTE.md` ergänzt.
- Sichtbare Einstiege für beide Prüfformen in `README.md` und `docs/DOKUMENTATION_ANWENDER.md` ergänzt.

## [3.6.1] - 2026-05-27

### Changed

- GUI-Polish: In der Aufgabenliste ist die Spalte `Confidence` jetzt ausgeblendet.
- Stabilität und Typisierung in `src/main.py` verbessert.
- Dokumentation in `docs/` und `memos/` bereinigt und aktualisiert.

### Quality

- Problems-Check ohne offene Diagnosen im Workspace.
- Regressionstest-Suite: 6/6 grün (`tools/test_regression_core.py`).

## [3.6.0] - 2026-05-27

### Added

- Zentrales Regelwerk in `data/config/import_rules.json` ausgebaut.
- Parser-Refactor in `src/word_processor.py`.
- Regressionstest- und QA-Rahmen mit Test-Suite, Matrix und Release-Checkliste.

### Quality

- Regressionstest-Suite: 6/6 grün.
- Parser-Mindesttests: 5/5 grün.
- Smoke-Test für Blockade- und Erfolgsfall verifiziert.

## [3.5.x] - 2025

### Changed

- Robuste Strukturübertragung, bessere Kompatibilität und stabilere Elementkopierung.
- Vollständige Aufgabenstruktur inkl. Überschriften und Formatierungen wird erhalten.
- Explizite Keyword-Extraktion und verbesserte Schwierigkeitsgrad-Erkennung eingeführt.
