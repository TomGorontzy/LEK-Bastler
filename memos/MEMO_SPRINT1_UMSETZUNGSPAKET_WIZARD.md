# MEMO: Sprint-1 Umsetzungspaket – Wizard für intuitive Aufgabenübernahme

## Ziel

Sprint 1 liefert einen robusten, geführten Import-Workflow, der typische Fehlimporte (v. a. Intro-Text als Aufgabe) vermeidet und vor dem Export eine verlässliche Vorschau bietet.

## Scope (Sprint 1)

- Wizard-Flow mit 4 Schritten in der GUI
- Parser-Qualität sichtbar machen (Confidence + Warnungen)
- Sicherheitsleitplanken (0 Aufgaben, leere Blöcke, unklare Struktur)
- „Vorschau == Export" sicherstellen

## Nicht-Scope (ab Sprint 2)

- Vollständige Regelkonfiguration via JSON
- Tiefer Parser-Refactor für Spezialfälle
- Umfangreiche Regressionstest-Suite

## Bestehende Codebasis (Ist-Analyse)

- GUI: `src/main.py`
- Extraktion/Erzeugung: `src/word_processor.py`
- Filterlogik: `src/task_selector.py`
- Vorlagen: `src/template_manager.py`

### Wichtige Beobachtungen

- `WordProcessor.extract_tasks()` nutzt bereits Struktur-Regeln (H1=Kategorie, H2=Aufgabe).
- Es fehlt eine explizite Validierungs-/Qualitätsstufe pro Aufgabe.
- In `word_processor.py` ist `_copy_elements_with_formatting` doppelt definiert (Cleanup nötig).
- `main.py` arbeitet aktuell direkt im Ein-Fenster-Flow; Wizard braucht eine kontrollierte Schrittlogik.

## Zielbild Architektur (Sprint 1)

## 1) Neues Modul: `src/import_wizard.py`

Enthält den Wizard-Zustand und die Schrittsteuerung.

### Datenmodell

```text
ImportTask
- id: int
- category: str
- title: str
- intro: list[str]              # optional
- content_elements: list        # bestehende all_elements
- difficulty: str
- keywords: list[str]
- confidence: str               # high | medium | low
- warnings: list[str]
- accepted: bool
```

```text
ImportSession
- source_file: str
- source_filename: str
- lek_theme: str
- tasks: list[ImportTask]
- global_warnings: list[str]
- approved_task_ids: set[int]
```

## 2) Erweiterung `WordProcessor`

Neue/angepasste Methoden:

- `extract_tasks_with_diagnostics(file_path) -> (tasks, report)`
  - baut auf `extract_tasks()` auf
  - ergänzt Diagnoseinfos pro Task:
    - leere Aufgabe
    - sehr kurzer Titel
    - Intro-dominiert
    - fehlende Keywords

- `_compute_confidence(task) -> str`
  - einfache Heuristik für Sprint 1:
    - high: Titel + Inhalt + keine Warnung
    - medium: leichte Warnungen
    - low: leer/unklar/strukturell auffällig

- `build_intro_and_body(task) -> (intro, body)`
  - trennt einleitende Abschnitte vor der Kernaufgabe für Vorschau

## 3) GUI-Anpassung in `main.py`

Neue Sektion „Import-Assistent" (Wizard):

1. Quelle wählen
2. Erkennung prüfen (Tabelle mit Confidence + Warnungen)
3. Vorschau (Ausgewählte Aufgaben)
4. Übernahme bestätigen

### UI-Elemente (minimal-invasiv)

- `ttk.Notebook` mit 4 Tabs oder Schrittpanel mit Weiter/Zurück
- Treeview-Spalten erweitern:
  - `Confidence`
  - `Warnungen`
- Checkbox/Toggle pro Aufgabe: übernehmen ja/nein
- Statusbereich oben: `Erkannt: X | Warnungen: Y | Low-Confidence: Z`

## 4) Vorschau==Export Garantie

Vor Export wird ausschließlich die freigegebene `approved_task_ids`-Menge exportiert.

- Keine Re-Extraktion zwischen Schritt 3 und 4.
- Export nutzt 1:1 die bereits bestätigten `content_elements`.

## Konkrete Umsetzungsschritte (Reihenfolge)

1. **Datenmodelle + Wizard-State**
   - Neues Modul `import_wizard.py`
   - Session-Objekt in GUI initialisieren

2. **Diagnostik in Extraktion integrieren**
   - `extract_tasks_with_diagnostics()` in `word_processor.py`
   - Confidence + Warnungen ergänzen

3. **GUI Schritt 2 (Erkennung prüfen)**
   - Treeview um Confidence/Warnungen erweitern
   - Selektionslogik für Übernahme

4. **GUI Schritt 3 (Vorschau)**
   - nur freigegebene Aufgaben anzeigen
   - Warn-Badge bei low confidence

5. **GUI Schritt 4 + Export-Kopplung**
   - Export nur bestätigter Aufgaben
   - Success-Dialog mit Zusammenfassung

6. **Cleanup**
   - doppelte Methode `_copy_elements_with_formatting` in `word_processor.py` bereinigen

## Akzeptanzkriterien (Sprint 1)

- Bei 0 erkannten Aufgaben erscheint ein klarer Hinweis mit Handlungsempfehlung.
- Jede Aufgabe zeigt Confidence (`high|medium|low`) und ggf. Warnungen.
- Nutzer kann Aufgaben vor Export explizit freigeben/abwählen.
- Export enthält nur freigegebene Aufgaben.
- Inhalt in Vorschau entspricht dem Export (keine stille Neuberechnung).
- Keine neuen Probleme in der Probleme-Konsole.

## Testfälle (Minimum)

1. Sauberes Quelldokument (nur H2-Aufgaben) → überwiegend `high`.
2. Dokument mit langem Intro vor erster Aufgabe → Intro-Warnung.
3. Dokument mit leerem Aufgabenblock → `low` + Warnung.
4. Nutzer wählt nur Teilmenge aus → Export enthält exakt diese.
5. Wizard-Abbruch und erneutes Laden → Session wird korrekt zurückgesetzt.

## Risiken + Gegenmaßnahmen

- Risiko: Heuristik zu streng/zu locker.
  - Gegenmaßnahme: Schwellenwerte zentral in Konstanten, leicht justierbar.
- Risiko: GUI wird zu komplex.
  - Gegenmaßnahme: Sprint 1 nur essentielle Elemente; Feinschliff später.
- Risiko: Code-Duplikate in `word_processor.py`.
  - Gegenmaßnahme: Cleanup-Schritt fest einplanen.

## Deliverables Sprint 1

- Neues Modul: `src/import_wizard.py`
- Erweiterungen in `src/main.py` und `src/word_processor.py`
- Stabiler Wizard-MVP mit Diagnose + sicherem Exportpfad
- Aktualisierte Doku in `memos/` und ggf. `docs/DOKUMENTATION_TECHNIK.md`

## Umsetzungsstand (2026-05-26)

### Bereits umgesetzt

- `src/import_wizard.py` eingeführt (`ImportTask`, `ImportSession`, Freigabe-Helfer)
- `WordProcessor.extract_tasks_with_diagnostics()` ergänzt
- Confidence-/Warnungsdaten in `main.py` eingebunden und in Treeview sichtbar gemacht
- Freigabe-Workflow ergänzt (`Auswahl freigeben`, `Freigaben löschen`)
- Exportlogik auf bestätigte Auswahl ausgerichtet
- Statuszeile + Export-Vorschau vor dem Speichern ergänzt
- Explizite Schrittführung im Wizard ergänzt (Schrittanzeige `1/4` bis `4/4` inkl. `Zurück/Weiter`)
- Duplikat-Cleanup in `word_processor.py` abgeschlossen (`_copy_elements_with_formatting` nur noch einmal definiert)

### Noch offen für Sprint-1-Abschluss

- Optional: kurze Technikdoku-Aktualisierung in `docs/DOKUMENTATION_TECHNIK.md`
- Abschluss-Smoke-Test mit 2–3 realen Quelldokumenten

## Smoke-Test Ergebnis (2026-05-26)

Automatisierter Testlauf (Python-Snippet in Projekt-Umgebung) mit realen Quelldokumenten:

- `Aufgaben_Auftragssteuerung und -koordination.docx` → `tasks=17`, `warnings=0`, `low=0`
- `Aufgaben_Mathematik-Grundlagen.docx` → `tasks=0`, `warnings=0`, `low=0`
- `Aufgaben_Personalwirtschaft.docx` → `tasks=0`, `warnings=0`, `low=0`

Zusätzliche Verifikation:

- ImportSession-Freigabe getestet (`session_approved=2`)
- Exportpfad getestet: `data/LEKs/_SMOKE_EXPORT.docx` erfolgreich erzeugt

Bewertung:

- Der neue Diagnose-/Freigabe-/Exportpfad funktioniert technisch stabil.
- Für zwei Quelldokumente wurden 0 Aufgaben erkannt; wahrscheinlich weicht die Struktur (z. B. Heading-Level) von der aktuellen H2-Regel ab und sollte im nächsten Schritt analysiert werden.

## Entscheidung Musterdatei (2026-05-26)

- Die Standard-Musterdatei `data/Vorlagen/AUFGABEN_MUSTER_STANDARD.docx` basiert nun auf
  `data/Vorlagen/AUFGABEN_GERUEST_WORD.docx`.
- Begründung: Die tabellenbasierte Struktur ist für die Erfassung robuster und wurde fachlich bevorzugt.
- Kurzvalidierung nach Umstellung: `TABLES=5`, `PARSED_TASKS=3`, `LOW=0`, `WARNINGS=1`.

## Hinweis zur Erkennungslogik (2026-05-26)

Die zwischenzeitliche Fallback-Nachschärfung für nicht standardisierte Sammlungen
(u. a. Mathematik/Personalwirtschaft) wurde wieder entfernt, da diese Dateien
fachlich als irrelevant eingestuft und aus `data/Aufgaben/` gelöscht wurden.

Aktiver Standard bleibt:

- Kategorie über `Heading 1`
- Aufgabe über `Heading 2`

Damit ist die Erkennung wieder bewusst auf die standardisierte Struktur ausgerichtet.

## Zwischenstand Strukturprüfung (2026-05-26, Auftragssteuerung exemplarisch)

Geprüfte Datei:

- `data/Aufgaben/Aufgaben_Auftragssteuerung und -koordination.docx`

Beobachtung:

- Die Datei nutzt eine tabellenbasierte Aufgabenstruktur (17 Tabellen, jeweils 2 Spalten Label/Wert),
  nicht ausschließlich klassische `Heading 1/Heading 2`-Absätze im Dokument-Body.
- Dadurch lieferte die reine H1/H2-Erkennung zunächst `0` Aufgaben.

Umsetzung:

- In `src/word_processor.py` wurde ein gezielter Fallback ergänzt:
  - Wenn H1/H2 keine Aufgaben ergibt, werden strukturierte Tabellen erkannt und geparst.
  - Unterstützte Kernfelder: `ID`, `Aufgabenstellung`, optional `Intro/Einleitung`,
   `Lösungsmöglichkeit/Hinweis`, `Schlagworte`, `Schwierigkeitsgrad`, `Kategorie`.

Validierung nach Anpassung:

- `Aufgaben_Auftragssteuerung und -koordination.docx` → `tasks=17`, `warning_task_count=0`, `low=0`
- `AUFGABEN_MUSTER_STANDARD.docx` → `tasks=3`, `warn=1`, `low=0`
- `AUFGABEN_GERUEST_WORD.docx` → `tasks=3`, `warn=1`, `low=0`

## Gemeinsame Review-Checkliste (Gliederung & Umsetzung)

Für den späteren gemeinsamen Durchgang der exemplarischen Datei:

1. **Kategorienlogik**
  - Sollen Kategorien explizit als Tabellenfeld gepflegt werden?
  - Oder weiterhin aus H1-Struktur außerhalb der Tabellen kommen?

2. **Titelstrategie**
  - Aktuell wird der Titel aus dem Beginn der Aufgabenstellung gebildet.
  - Soll es ein eigenes Feld `Titel` geben (stabiler für Übersicht/Filter)?

3. **Schwierigkeitsgrad-Regel**
  - Werte wie `leicht | mittel | schwer`: gewünschte Interpretation klären
    (ein Wert, Mehrfachauswahl oder Standardwert).

4. **Pflichtfelder im Template**
  - Minimal verpflichtend: `ID`, `Aufgabenstellung`
  - Optional: Intro, Hinweis, Schlagworte, Schwierigkeit, Kategorie

5. **Exportdarstellung**
  - Prüfen, ob Intro/Hinweis im finalen LEK-Layout genau wie gewünscht erscheinen.

6. **Governance für neue Sammlungen**
  - Kurzregel in Doku: „Entweder H1/H2-Struktur oder Tabellen-GERUEST nach Feldschema“.

## Entscheidungsgrundlage für den gemeinsamen Review (Stand 2026-05-26)

Aktuelle Messwerte aus `Aufgaben_Auftragssteuerung und -koordination.docx`:

- Import: `17` Aufgaben erkannt
- Diagnostik: `warning_task_count=0`, `low_confidence_count=0`
- IDs: `17/17` eindeutig, `0` fehlend
- Kategorien: aktuell `17x Ohne Kategorie` (kein Tabellenfeld `Kategorie` befüllt)
- Schwierigkeit (roh):
  - `leicht`: 8
  - `mittel`: 6
  - `schwer`: 2
  - `leicht | mittel | schwer`: 1 (mehrdeutig)

Empfohlene Festlegungen (für Team-Entscheid):

1. **Kategorie als Pflichtfeld im Tabellenmodus**
  - Empfehlung: `Kategorie` verpflichtend pflegen.
  - Fallback ohne Kategorie bleibt technisch erlaubt, wird aber als Doku-Verstoß gewertet.

2. **Mehrdeutige Schwierigkeit normalisieren**
  - Empfehlung: genau ein Zielwert pro Aufgabe (`leicht|mittel|schwer`).
  - Falls mehrere Werte eingetragen sind, Standard auf `Mittel` + Warnhinweis im Import.

3. **Titel-Feld perspektivisch ergänzen**
  - Kurzfristig: Titel aus Aufgabenstellung (funktional ausreichend).
  - Mittelfristig: eigenes Tabellenfeld `Titel` für bessere Lesbarkeit/Filter.

4. **Feldschema für neue Sammlungen verbindlich machen**
  - Pflicht: `ID`, `Aufgabenstellung`, `Schwierigkeitsgrad`, `Kategorie`
  - Optional: `Intro/Einleitung`, `Lösungsmöglichkeit/Hinweis`, `Schlagworte`

## Update nach Anwender-Feedback (2026-05-26)

- `AUFGABEN_MUSTER_STANDARD.docx` und
  `Aufgaben_Auftragssteuerung und -koordination.docx` wurden um das Tabellenfeld `Kategorie` ergänzt.
- Aktuelle Validierung:
  - `AUFGABEN_MUSTER_STANDARD.docx` → Kategorien gesetzt (`Aufgaben-Gerüst (Word)`)
  - `Aufgaben_Auftragssteuerung und -koordination.docx` → Kategorien gesetzt (`Auftragssteuerung und -koordination`)

Zusätzlich umgesetzt (Verbindliche Qualitätssicherung vor Export):

- Bei erkannten Inkonsistenzen im Schwierigkeitsgrad (z. B. mehrere Werte in einem Feld)
  wird eine explizite Warnung erzeugt.
- Solche Inkonsistenzen blockieren den Export in der GUI, bis die Quelle durch den
  jeweiligen Anwender bereinigt und neu geladen wurde.

## Erweiterung: Neue Aufgabe per EXE-GUI übernehmen (2026-05-26)

Neu in der GUI (`src/main.py`):

- Button: `Aufgabe aus Word übernehmen...`
- Ablauf:
  1. Ziel-Aufgabensammlung ist die aktuell geladene Datei.
  2. Separate Quell-Datei auswählen (neue Aufgabe).
  3. Metadaten erfassen (Kategorie, Schwierigkeitsgrad, Schlagworte).
  4. Vorschau-/Bestätigungsdialog zeigt Ziel-ID, Metadaten und Quellstruktur (Absätze/Tabellen).
  5. Aufgabe wird erst nach Bestätigung als neue Tabellenaufgabe in die Sammlung übernommen.

Optionale Detailansicht vor Bestätigung:

- Vor der finalen Übernahme kann der Anwender `Details anzeigen` wählen.
- Die Ansicht listet die ersten Quellblöcke auf (Absätze + Tabellen inkl. Tabellenabmessungen/erste Zelle),
  um die inhaltliche Plausibilität vor dem Schreiben zu prüfen.

Sicherheitsnetz bei Übernahme:

- Vor dem tatsächlichen Schreiben wird automatisch eine zeitgestempelte Backup-Kopie
  der Ziel-Aufgabensammlung erstellt.
- Der Pfad zur Backup-Datei wird im Erfolgsdialog angezeigt.

Bulk-Übernahme (mehrere Quell-Dateien):

- Neuer GUI-Button: `Mehrere Aufgaben aus Word übernehmen...`
- Einmalige Metadatenabfrage (Kategorie, Schwierigkeit, Schlagworte) für die Serie.
- Für jede Quell-Datei: Preview + optionale Detailansicht + Entscheidung
  `übernehmen / überspringen / Serie abbrechen`.
- Optional können pro Datei die Metadaten (Kategorie, Schwierigkeit, Schlagworte)
  vor der Übernahme individuell überschrieben werden.
- Abschlussdialog mit Serienzusammenfassung (`übernommen`, `übersprungen`, `Fehler`).

Duplikat-Check vor Import:

- In der Import-Vorschau wird ein Duplikatverdacht per Textähnlichkeitsprüfung ermittelt.
- Bei Treffer zeigt die GUI den wahrscheinlichsten Match (ID, Titel, Kategorie, Ähnlichkeit).
- Entscheidung ist explizit je Datei:
  - Einzelimport: `trotzdem fortfahren?`
  - Bulk-Import: `fortfahren / überspringen / Serie stoppen`

Regelkonfiguration (Sprint-Fortsetzung):

- Neue Konfigurationsdatei: `data/config/import_rules.json`
- Aktuell konfigurierbar:
  - `duplicate_mode` (`strict|normal|relaxed`)
  - `duplicate_similarity_thresholds` (Modusabhängige Schwellwerte)
  - `duplicate_similarity_threshold` (Fallback-Schwellwert)
  - `max_preview_blocks` (Anzahl Detailblöcke in der Vorschau)
  - `bulk_max_errors` (Abbruchgrenze bei Fehlern in Serienübernahme)
  - `category_rules` (`required`, `block_export_on_missing`, `missing_values`)
  - `default_import_metadata` (`category`, `difficulty`, `keywords`)
- Ziel: Feinjustierung ohne Codeänderung für künftige Sprint-Iterationen.

GUI-Erweiterung (Laufzeit-Preset):

- In `src/main.py` wurde ein Preset-Umschalter für die Duplikaterkennung ergänzt.
- Anwender können in der EXE/GUI zwischen `strict`, `normal`, `relaxed` wechseln.
- Die aktive Schwellwertzahl wird direkt im UI angezeigt (z. B. `0.82` bei `strict`).
- Der Modus wird für die laufende Sitzung angewendet (ohne Dateiedit nötig).
- Zusätzlich kann der gewählte Modus per Button `Als Standard speichern` dauerhaft in
  `data/config/import_rules.json` geschrieben werden.

Zusätzliche Validierung vor Übernahme:

- Inkonsistenter Schwierigkeitsgrad in der Eingabe (z. B. mehrere Werte wie `leicht | mittel`)
  blockiert die Übernahme mit klarem Korrekturhinweis.
- Kategorie ist als Eingabe verpflichtend (leer wird abgewiesen).
- Schwierigkeitsgrad wird auf `leicht|mittel|schwer` normalisiert (inkl. Alias-Mapping wie `easy`, `normal`, `hard`).

Kategorienlogik (verbindlich umgesetzt):

- Kategorie ist als Pflichtfeld konfiguriert (`category_rules.required=true`).
- Fehlende/ungültige Kategorien (z. B. `Ohne Kategorie`, leer) werden diagnostisch markiert.
- Wenn `category_rules.block_export_on_missing=true`, blockiert der Export betroffene Aufgaben,
  bis die Kategorie in der Quelle korrigiert und neu geladen wurde.

Technische Umsetzung (`src/word_processor.py`):

- Neue Methode `append_task_from_document(...)` ergänzt.
- Inhalt der Quelldatei wird in das Feld `Aufgabenstellung (Pflicht)` der neuen Aufgabe übernommen.
- Unterstützt rich content (Absätze, Tabellen, Formatierungen; eingebettete Inhalte best effort).
- Nach Übernahme wird die Sammlung gespeichert und im GUI neu geladen.
