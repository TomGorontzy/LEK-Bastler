# RELEASE QA Checkliste (Sprint 4)

## Ziel

Kurzer, reproduzierbarer Pflichtdurchlauf vor jedem Release.

## 1) Automatisierte Kernprüfung

- Regressionstests ausführen (`tools/test_regression_core.py`)
- Erwartung: alle Tests grün

## 2) Smoke-Test (fachlicher Ablauf)

1. Aufgabensammlung laden
2. Diagnosewerte prüfen (Warnungen/Confidence plausibel)
3. Teilmenge freigeben
4. Schritt-3-Gesamtvorschau prüfen
5. Export durchführen
6. Ergebnisdokument fachlich gegenprüfen

## 2a) Post-Release-Verifikation (skriptbar)

- ZIP-Artefakt vorhanden, Größe + SHA256 protokolliert
- ZIP in isoliertes Temp-Verzeichnis entpackt
- Pflichtinhalte vorhanden (`*.exe`, `README.md`, `LIZENZ.txt`, `data/`, `docs/`, `data/LEKs/README.md`)
- EXE startet bis GUI-Idle (kurzer Starttest), danach kontrolliert beendet
- Ergebnisse im Release-Smoketest-Protokoll dokumentiert

## 3) Sollkriterien

- Vorschau == Export-Reihenfolge für strukturierte Aufgaben
- Blockierregeln greifen (Kategorie/Pflichtfelder/Difficulty)
- Teilfreigabe exportiert exakt freigegebene Aufgaben
- Keine neuen Fehler in der Probleme-Konsole

## 4) Freigabeentscheidung

Release nur freigeben, wenn:

- automatisierte Regression grün
- Smoke-Checkliste vollständig durchlaufen
- Post-Release-Verifikation inkl. Protokoll abgeschlossen
- keine offenen Blocker in Kernpfaden
