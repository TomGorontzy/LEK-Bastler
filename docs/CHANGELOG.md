# Changelog

Alle relevanten Änderungen an diesem Projekt werden hier dokumentiert.

## [Unreleased]

### Added

- Keine Änderungen dokumentiert.

### Changed

- Keine Änderungen dokumentiert.

## [3.7.6] - 2026-06-07

### Added

- LEK-Export erzeugt jetzt zusätzlich automatisch eine Protokolldatei (`.txt`) mit gleichem Basisnamen wie die exportierte LEK-Datei.
- Die Protokolldatei enthält die im Export verwendeten Aufgabennummern in der tatsächlichen Exportreihenfolge.

### Changed

- GUI-Workflow für Freigabe/Export verbessert:
  - Nach Klick auf `Auswahl freigeben` werden die Exportaktionen (`Auswahl exportieren`, `Freigaben exportieren`) sofort aktiviert.
  - Die Aktivierung ist nicht mehr unnötig an den Wizard-Schritt 4 gebunden, sobald freigegebene Aufgaben vorliegen.
- Aktionslayout verfeinert:
  - Die Export-Gruppe wird jetzt oben rechts ausgerichtet (statt vertikal zentriert), für konsistentere Bedienführung.

### Quality

- Regressionstest-Suite erfolgreich: **33/33 OK** (`tools/test_regression_core.py`).

## [3.7.5] - 2026-06-07

### Fixed

- Externe Tabellenreferenzen im strukturierten Export korrigiert:
  - Marker-Kombinationen wie `<<tabelle=...>>` und `<<tabelle_format=...>>` führen nicht mehr zu doppeltem Einfügen derselben externen Datei.
  - Der Exportpfad verhindert jetzt explizit den Fallback-Doppelinsert, wenn die externe Tabelle bereits beim Marker-Parsing übernommen wurde.
- Stil-/Schriftübernahme bei extern referenzierten Dokumenten gehärtet:
  - Konflikte bei gleichnamigen Style-IDs (z. B. `Normal`) werden konfliktfrei remappt (`ext_*`).
  - Implizite Default-Absatzstile externer Inhalte werden materialisiert und auf remappte Styles umgebogen, sodass Schriftkontexte stabil übernommen werden.
- GUI-Layout stabilisiert:
  - Flattern im Bereich unterhalb des Import-Assistenten bei Resize wurde durch Guard-Logik gegen redundante Re-Layouts beseitigt.

### Quality

- Regressionstest-Suite erweitert und erfolgreich ausgeführt: **32/32 OK** (`tools/test_regression_core.py`).

## [3.7.4] - 2026-06-07

### Added

- Verbindliche Markdown-Lintprüfung im Build (`src/build.ps1`):
  - Der Build führt jetzt standardmäßig `markdownlint-cli2` für `README.md`, `docs/**/*.md`, `release/**/*.md` und `data/**/*.md` aus.
  - Bei Lint-Verstößen wird der Build mit einer klaren Fehlermeldung abgebrochen.
  - Optional kann der Lint-Schritt explizit über `-SkipLint` übersprungen werden.

### Changed

- Build-Dateischreiben in `src/build.ps1` auf robustes UTF-8-Write (`WriteAllText`) umgestellt, um unnötige zusätzliche Leerzeilen in Markdown-Dateien zu vermeiden.
- GUI-Aktionsbereich in `src/main.py` neu layoutet:
  - Gruppen `Auswahl` und `Prüfung` links untereinander,
  - Gruppe `Export` rechtsbündig,
  - adaptive Gruppenbreite bei Fenstergrößenänderungen für bessere Nutzbarkeit auf kleineren Displays.

### Fixed

- Nummerierungsprüfung ausgebaut: Doppelte Aufgabennummern werden beim Laden erkannt, als blockierende Warnung geführt und vor Export klar eingefordert zu bereinigen.
- Mehrere Markdown-Regelverstöße in Projekt- und Release-Dokumenten bereinigt (u. a. MD012/MD022/MD041), damit Build-Lint stabil grün bleibt.

### Quality

- Regressionstest-Suite erfolgreich: **30/30 OK** (`tools/test_regression_core.py`).
- Build-Lintprüfung erfolgreich: `markdownlint-cli2` meldet **0 Fehler** im Buildlauf (`src/build.ps1 -SkipBuild`).

## [3.7.3] - 2026-06-07

### Changed

- LEK-Exportlayout für Aufgabenabstände präzisiert:
  - Zwischen `Überschrift 1` (Aufgabentitel) und Aufgabeninhalt wird jetzt konsistent nur **eine Leerzeile** ausgegeben.
  - Zwischen zwei Aufgaben werden jetzt **maximal zwei Leerzeilen** ausgegeben.
- Export für Aufgaben mit Objekt-/Tabellenstart stabilisiert:
  - Doppelte Leerzeilen vor einem ersten Tabellen-/Objektblock wurden entfernt.
  - Führende und abschließende leere Absätze innerhalb eines Aufgabenblocks werden robuster begrenzt.
- Nummerierte Listen im LEK-Export korrigiert:
  - Nummerierte Absätze in Aufgabenbeschreibungen behalten jetzt ihre lokale Nummerierung pro Aufgabe, statt ungewollt in späteren Aufgaben weiterzulaufen.
  - Das gilt auch für nummerierte Listen in verschachtelten Tabellen innerhalb strukturierter Aufgaben.

### Quality

- Syntax-/Problems-Check für `src/word_processor.py` ohne Fehler.
- Mini-End-to-End-Exporttest (inkl. tabellenstartender Aufgabe) bestätigt die Zielabstände.
- Regressionstest-Suite erfolgreich: **27/27 OK** (`tools/test_regression_core.py`).

## [3.7.2] - 2026-06-07

### Changed

- GUI-Titelleiste/EXE-Versionserkennung stabilisiert: Wenn `build_version_info.txt` im Deploy fehlt, wird die Versionsnummer robust aus EXE-/Ordnernamen (z. B. `LEK-Bastler_3.7.0.exe`) ermittelt und wieder im Fenstertitel angezeigt.
- Auswahl-, Freigabe- und Exportlogik für Aufgabenfamilien erweitert: Hauptaufgaben und zugehörige Unteraufgaben (z. B. `1.0`, `1.1`, `1.2`) werden jetzt automatisch gemeinsam behandelt und nur als vollständige Gruppe in die LEK übernommen.
- Externe Großtabellen können jetzt per Marker `<<tabelle=Dateiname>>` aus einem lernbereichsspezifischen Unterordner unter `data/Aufgaben/` referenziert und beim Export automatisch in passender Hoch-/Querformat-Orientierung übernommen werden.
- Übernahme leerer Absätze beim Aufgaben-Export weiter gehärtet:
  - Leere Zeilen bleiben nach Aufgaben-Bearbeitung (`Mehrfach-/Einzelbearbeitung`) im Inhaltsmodell erhalten.
  - Fallback-Export schreibt leere Inhaltszeilen jetzt explizit als leere Word-Absätze.

## [3.7.1] - 2026-06-06

### Changed

- Diagramm-Dokumentation für Mermaid-Workflows erweitert und stabilisiert:
  - Starter-Snippet und Mini-Do/Don't-Liste für renderstabile Styles ergänzt.
  - Hinweis auf mermaid-cli-kompatible `linkStyle`-Nutzung präzisiert.
- Typrobustheit in der GUI-Logik verbessert:
  - Guard für Blocking-Summary in `src/main.py` ergänzt.
  - Ergebnis-Typisierung im Metadaten-Dialog präzisiert, um Pylance-Fehler zu vermeiden.

### Quality

- Regressionstest-Suite erfolgreich: **23/23 OK** (`tools/test_regression_core.py`).

## [3.7.0] - 2026-06-06

### Added

- Wizard-/Listen-UX ausgebaut:
  - **Nur Blocker anzeigen** und **Nur Blocker auswählen** ergänzt.
  - Aufgabenliste nach Spalten sortierbar.
  - Doppelklick-Bearbeitung einzelner Aufgaben ergänzt.
  - Mehrfachbearbeitung erweitert (Kategorie/Schwierigkeit, Schlagworte ersetzen oder anhängen).
- Aufgabenanlage und Metadatenpflege verbessert:
  - Eingabehilfe ergänzt (Button **"Eingabehilfe (JSON & Formulierungen)"**) samt konfigurierbarer Datei `data/config/task_authoring_hints.json`.
  - Strukturierter Metadaten-Dialog mit Live-Validierung und direktem Hilfezugriff ergänzt.
  - Inline-Feldhilfen (`field_hints`) und Smart Defaults pro Kategorie (`category_defaults`) ergänzt.
  - Kategorieeingabe als editierbare Autovervollständigung ergänzt; Vorschläge werden nach Nutzung priorisiert.
  - Zuletzt verwendete Kategorie wird benutzerspezifisch in `data/config/task_authoring_user_config.json` gespeichert.
- Modussteuerung in der GUI ergänzt:
  - Neuer Modus-Schalter `einfach` vs. `erweitert`.
  - Erweiterte Aktionen (Word-Import, Duplikat-Preset, Blockerfilter) nur im Modus `erweitert`.
  - Erweiterte Funktionen werden im Modus `einfach` vollständig ausgeblendet; Anzeige `einfach (empfohlen)` ergänzt.
  - Auswahlkriterium **Max. Anzahl Aufgaben** auf 20 begrenzt (Default 20).
- Export und Layout erweitert:
  - Übernahme von Tabellen und Inhaltssteuerelementen (z. B. Kontrollkästchen) beim LEK-Export.
  - Aufgabentitel als `Überschrift 1`; Punkte am Titelende rechtsbündig mit Rahmen.
  - Punkte-Box visuell vereinheitlicht und über `data/config/import_rules.json` (`export_rules.title_points_box.*`) konfigurierbar gemacht.
  - GUI-Eingabehilfe verweist explizit auf die Export-Layout-Keys.
- Versions- und Vorlagenpflege:
  - GUI-Titelleiste zeigt die Version aus `build_version_info.txt`.
  - Vorlagenbereinigung: `AUFGABEN_MUSTER_STANDARD.docx` → `AUFGABEN-Vorlage.docx`; `AUFGABEN_GERUEST_WORD.docx` entfernt.

### Changed

- Wizard-Status zeigt neben Warnungs-/Blockeranzahl jetzt konkrete **To-fix-Hinweise** (z. B. Titel/Kategorie/Schwierigkeit).
- Sortierzustand wird in der Aufgabenliste visuell über Header-Pfeile (`▲`/`▼`) angezeigt.

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
