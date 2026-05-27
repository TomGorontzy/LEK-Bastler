# MEMO: Sprint-3 Umsetzungspaket – Parser-Refactor und Spezialfälle

## Ziel

Sprint 3 strukturiert die Aufgaben-Erkennung und Metadatenextraktion neu, damit Sonderfälle nicht mehr primär über punktuelle Fallbacks, sondern über klar getrennte Parser-Bausteine behandelt werden.

## Scope (Sprint 3)

- Parser-Refactor in `src/word_processor.py`
- Saubere Trennung von Erkennungsmodi (H1/H2 vs. Tabellenmodus)
- Entkopplung von Struktur-Erkennung, Metadatenextraktion und Diagnose
- Robusteres Handling gemischter Dokumentstrukturen

## Nicht-Scope (ab Sprint 4)

- Umfangreicher Fixture-Katalog für Regressionstests
- Vollständige Performance-Optimierung großer Dokumente
- Größere UX-Erweiterungen im Wizard

## Ausgangslage

Aktuell ist die Verarbeitung funktional, aber mehrere Verantwortlichkeiten liegen dicht beieinander:

- Aufgaben erkennen
- Intro-Kontext erkennen
- Tabellenfelder normalisieren
- Diagnosewarnungen berechnen
- Exportvorschau vorbereiten

Die aktuelle Lösung ist stabil genug für Sprint 1, sollte aber vor wachsender Regellast modularisiert werden.

## Zielbild Architektur (Sprint 3)

## 1) Parser-Pipeline in Schichten

### Schicht A: Strukturerkennung

- H1/H2-basierte Aufgabenquellen
- strukturierte Tabellenquellen
- ggf. gemischte Quellen mit klarer Priorität

### Schicht B: Normalisierung

- IDs / sichtbare Nummern
- Titel / Kategorie / Difficulty / Keywords
- Feldalias-Auflösung

### Schicht C: Diagnose

- Confidence
- Warnungen
- Intro-/Kontextmarkierungen
- Exportblocker

### Schicht D: Vorschau-/Exportvorbereitung

- Fließtextreihenfolge
- Delta-Check
- Preview-Linien / Exportabschnitte

## 2) Zielmethoden / Verantwortungen

Beispiele für zu separierende Verantwortungsblöcke:

- `_extract_tasks_from_headings(...)`
- `_extract_tasks_from_structured_tables(...)`
- `_normalize_task_fields(...)`
- `_build_task_diagnostic(...)`
- `build_task_flow_preview_lines(...)`

## Konkrete Umsetzungsschritte (Reihenfolge)

1. **Parser-Verantwortungen kartieren**
   - Bestehende Methoden clustern
   - Dopplungen/Seiteneffekte identifizieren

2. **Erkennungsmodi sauber trennen**
   - H1/H2-Modus
   - Tabellenmodus
   - Prioritäts- und Fallback-Regeln dokumentieren

3. **Normalisierung auslagern**
   - Nummernlogik
   - Intro-/Hint-/Title-Felder
   - Difficulty-/Keyword-Normalisierung

4. **Diagnosepfad vereinheitlichen**
   - Confidence nur an einer Stelle berechnen
   - Warnungen deterministisch sammeln

5. **Exportnahe Vorschau andocken**
   - Keine stille Sonderlogik im GUI-Code
   - Preview-/Export-Bausteine aus denselben Daten ableiten

## Akzeptanzkriterien (Sprint 3)

- Parser-Pfade sind fachlich klar getrennt.
- Tabellen- und H1/H2-Modus lassen sich unabhängig nachvollziehen.
- Metadaten-Normalisierung ist an zentralen Stellen gebündelt.
- Vorschau- und Exportvorbereitung bauen auf denselben normalisierten Daten auf.
- Keine neuen Probleme in der Probleme-Konsole.

## Testfälle (Minimum)

1. Reines H1/H2-Dokument → Aufgaben korrekt erkannt.
2. Reines Tabellen-Dokument → strukturierte Aufgaben korrekt erkannt.
3. Dokument mit Intro-/Kontextblock → Intro-Markierung stabil.
4. Gemischtes Dokument → definierte Priorität/Fallback greift nachvollziehbar.
5. Titel-/Kategorie-/Difficulty-Normalisierung bleibt unverändert korrekt.

## Risiken + Gegenmaßnahmen

- Risiko: Refactor verschiebt unbemerkt bestehendes Verhalten.
  - Gegenmaßnahme: Vorher/Nachher-Snippets mit echten Beispieldokumenten.
- Risiko: Zu große Änderung in einem Schritt.
  - Gegenmaßnahme: Teilrefactor pro Verantwortungsblock.
- Risiko: Exportlogik koppelt weiter implizit an Rohdaten.
  - Gegenmaßnahme: Export nur über normalisierte Task-Daten vorbereiten.

## Deliverables Sprint 3

- Refaktorierte Parser-Struktur in `src/word_processor.py`
- Dokumentierte Erkennungsmodi und Fallback-Regeln
- Stabilere Trennung von Erkennung, Normalisierung und Diagnose
- Aktualisierte Technikdoku für Parser-/Verarbeitungspfad
