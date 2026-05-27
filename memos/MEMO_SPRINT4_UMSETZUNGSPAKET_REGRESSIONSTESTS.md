# MEMO: Sprint-4 Umsetzungspaket – Regressionstests und Qualitätssicherung

## Ziel

Sprint 4 sichert die in Sprint 1 bis 3 erreichte Funktionalität durch eine systematische Regressionstest- und Smoke-Test-Strategie ab. Ziel ist ein verlässlicher Release-Prozess mit reproduzierbaren Prüfungen.

## Scope (Sprint 4)

- Regressionstest-Konzept für Kernpfade
- Testdaten/Fixures aus realistischen Beispielen
- Smoke-Test-Checkliste für Releases
- Fokus auf Vorschau==Export, strukturierte Tabellen und Validierungsregeln

## Nicht-Scope (später)

- Vollautomatische End-to-End-GUI-Automation
- Performance-Benchmarks großer Dokumentmengen
- CI/CD-Ausbau über den lokalen Qualitätsrahmen hinaus

## Zu sichernde Kernbereiche

- H1/H2-Erkennung
- Tabellenmodus mit Feldaliasen
- Titel-Fallback
- Kategoriepflicht
- Difficulty-Normalisierung und Exportblockade
- Nummerndarstellung (`1`, `1.0`, `1.1`, `1.2`)
- Fließtext-Export strukturierter Aufgaben
- Schritt-3-Gesamtvorschau und Delta-Check

## Zielbild Teststrategie

## 1) Regressionstest-Katalog

Mindestens ein Testfall je Kernverhalten:

- sauberes Standarddokument
- Dokument mit Intro/Einleitung
- strukturierte Tabelle mit optionalen Feldern
- Tabelle ohne optionale Felder
- fehlende Kategorie
- fehlender Titel mit Fallback
- inkonsistente Schwierigkeit
- selektiver Export nur freigegebener Aufgaben

## 2) Smoke-Test-Set für Releases

Kompakter manueller bzw. skriptbarer Durchlauf vor jedem Release:

1. Datei laden
2. Diagnosewerte prüfen
3. Teilmenge freigeben
4. Schritt-3-Vorschau prüfen
5. Export durchführen
6. Ergebnisdokument fachlich plausibel gegenprüfen

## 3) Erwartete Artefakte

- definierte Testquellen / Musterdateien
- dokumentierte Soll-Ergebnisse
- reproduzierbare Prüfschritte pro Release

## Konkrete Umsetzungsschritte (Reihenfolge)

1. **Kritische Verhaltensmatrix erstellen**
   - Welche Funktionen müssen stabil bleiben?
   - Welche Dokumenttypen decken diese ab?

2. **Fixture-/Beispieldokumente auswählen**
   - vorhandene reale Beispiele
   - ggf. reduzierte Testdokumente ableiten

3. **Smoke-Test-Checkliste standardisieren**
   - Laden
   - Freigeben
   - Vorschau
   - Export
   - Ergebnisprüfung

4. **Regressionserwartungen festhalten**
   - Task-Zahl
   - Warnungen / Low-Confidence
   - Exportblockaden
   - sichtbare Nummern
   - enthaltene/fehlende Blöcke im Export

5. **Release-Prozess koppeln**
   - Prüfen, welche Mindesttests vor Tag/Release verpflichtend sind

## Akzeptanzkriterien (Sprint 4)

- Es existiert eine nachvollziehbare Regressionstest-Matrix.
- Kernverhalten ist für reale Dokumente reproduzierbar abgesichert.
- Vorschau==Export ist für die relevanten Fälle explizit geprüft.
- Release-Smoketests sind klar dokumentiert und wiederholbar.
- Keine neuen Probleme in der Probleme-Konsole.

## Testfälle (Minimum)

1. Tabellenaufgabe mit Intro/Hinweis/Punkten → vollständige Vorschau und Exportreihenfolge.
2. Tabellenaufgabe ohne optionale Blöcke → Delta-Check meldet Lücken, Export bleibt stabil.
3. Aufgabe mit fehlender Kategorie → Diagnostik + Exportblockade.
4. Aufgabe mit fehlendem Titel → Fallback greift, Warnung nachvollziehbar.
5. Inkonsistente Schwierigkeit → Warnung + Blockade.
6. Teilfreigabe im Wizard → Export enthält exakt die freigegebenen Aufgaben.

## Risiken + Gegenmaßnahmen

- Risiko: Realdokumente ändern sich und machen Testergebnisse instabil.
  - Gegenmaßnahme: reduzierte stabile Testquellen/Fixures ergänzen.
- Risiko: Smoke-Tests werden im Alltag übersprungen.
  - Gegenmaßnahme: kurze Pflichtcheckliste vor jedem Release.
- Risiko: Vorschau- und Exportprüfung bleibt zu subjektiv.
  - Gegenmaßnahme: Sollkriterien je Testfall definieren.

## Deliverables Sprint 4

- Regressionstest-Matrix in `memos/` oder `docs/`
- Standardisierte Smoke-Test-Checkliste
- Dokumentierte Release-QA-Mindestschritte
- Abgesicherte Kernpfade für künftige Releases
