# MEMO: Regressionstest-Matrix (Sprint 4)

## Zweck

Diese Matrix definiert die stabil zu haltenden Kernverhalten und deren Soll-Ergebnisse.

## Matrix

| ID | Bereich | Eingabe/Fixture | Erwartung |
|---|---|---|---|
| R1 | Strukturierte Tabelle (voll) | Tabelle mit Titel, Intro, Aufgabenstellung, Hinweis, Punkte | Vorschau/Export-Reihenfolge: Titel → Intro → Aufgabenstellung → Hinweis → Punkte |
| R2 | Strukturierte Tabelle (minimal) | Tabelle ohne Intro/Hinweis/Punkte | Delta-Check meldet fehlende optionale Blöcke, Export bleibt stabil |
| R3 | Kategoriepflicht | Tabelle mit leerer Kategorie | Warnung „Kategorie fehlt“, Exportblockade-Regel ist aktiv |
| R4 | Titel-Fallback | Tabelle ohne Titel, mit Aufgabenstellung | Titel wird aus Aufgabenstellung abgeleitet; Pflichtfeldwarnung bleibt nachvollziehbar |
| R5 | Difficulty-Konsistenz | Schwierigkeit mit mehreren Werten | Warnung „Inkonsistenter Schwierigkeitsgrad“, Blockade-Regel aktiv |
| R6 | Teilfreigabe Wizard | 3 Aufgaben, davon 2 freigegeben | Export enthält exakt freigegebene Aufgaben |

## Mindestabdeckung pro Release

- Alle Fälle R1–R6 müssen vor Release einmal erfolgreich geprüft werden.
- Bei Parser-/Regeländerungen zusätzlich Smoke-Test mit realem Quelldokument.
