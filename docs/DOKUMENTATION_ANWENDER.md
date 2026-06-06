# DOKUMENTATION_ANWENDER

## 1. Zweck

Mit dem LEK-Bastler erstellen Sie aus vorbereiteten Aufgaben-Worddateien schnell eine Lernerfolgskontrolle (LEK).

## 2. Start

### Eigenständige Windows-EXE

1. Öffnen Sie den ausgelieferten Ordner `LEK-Bastler_<Version>`.
2. Starten Sie `LEK-Bastler_<Version>.exe`.

### Python-Ausführung (Entwicklung)

1. Python-Umgebung vorbereiten
2. Anwendung starten über `src/main.py`

Hinweis zum aktuellen Stand: Version **3.7.0**.

## 3. Typischer Ablauf

1. **Aufgabensammlung laden** (`data/Aufgaben/*.docx`)
2. **Aufgaben filtern** (optional, z. B. über Suchbegriffe/Schwierigkeit)
3. **Aufgaben auswählen**
4. **Export starten**
5. Zieldatei im Ordner `data/LEKs/` speichern

Die vorgeschlagenen Dateinamen werden automatisch erzeugt und können bei Bedarf geändert werden.

## 3a. Visualisierter Ablauf

![Anwenderablauf LEK-Bastler](diagramme/anwender_ablauf.svg)

Das Diagramm zeigt den typischen Pfad von der Aufgabensammlung bis zum Export in `data/LEKs/`.

## 4. Wichtige Hinweise

- Der Ordner `data/LEKs/` wird beim Export automatisch erstellt, falls er fehlt.
- Für beste Ergebnisse sollten Aufgaben sauber mit Überschriften oder als strukturierte 2-Spalten-Tabellen aufgebaut sein.
- Vorlagen liegen im Ordner `data/Vorlagen/`.
- Relevante Vorlagen im aktuellen Stand:
  - `data/Vorlagen/LEK-Vorlage.docx` (LEK-Export)
  - `data/Vorlagen/AUFGABEN-Vorlage.docx` (Muster für Aufgabensammlungen)
- Falls keine passende Spezialvorlage gefunden wird, wird die Standardvorlage verwendet.
- Vorschau und Export nutzen dieselbe fachliche Reihenfolge bei strukturierten Aufgaben (Titel → Intro → Aufgabenstellung → Hinweis → Punkte).

## 4a. Regelwerk im Alltag (ab v3.6.0)

- Viele Import-/Validierungsregeln sind zentral in `data/config/import_rules.json` konfigurierbar.
- Änderungen an Aliasen oder Pflichtfeldern können in der Regel ohne Codeänderung erfolgen.
- Bei Regelverstößen (z. B. fehlende Kategorie, inkonsistente Schwierigkeit) kann der Export bewusst blockiert werden.

## 5. Fehlerbehebung

### App startet nicht

- Prüfen, ob alle Dateien des Windows-Releases vorhanden sind.
- EXE einmal lokal auf Laufwerk statt direkt aus Cloud-Sync-Cache starten.
- Prüfen, ob Virenschutz/SmartScreen das Starten blockiert.

### Export schlägt fehl

- Sicherstellen, dass die Eingabedatei nicht in Word gesperrt ist.
- Schreibrechte im Zielordner prüfen.
- Prüfen, ob `data/Vorlagen/LEK-Vorlage.docx` vorhanden ist.
- Bei ungewöhnlichen Formatproblemen mit einer alternativen Aufgabenquelle gegenprüfen.

### Icon wird nicht korrekt angezeigt

- Neue EXE nach Build verwenden.
- Bei Bedarf Windows Explorer neu starten (Iconcache-Aktualisierung).

### data/LEKs-Ordner fehlt im Windows-Release

- In aktuellen Releases ist `data/LEKs/` enthalten.
- Beim Export wird `data/LEKs/` zusätzlich zur Laufzeit angelegt, falls nötig.

## 6. Datenablage

- Eingaben: `data/Aufgaben/`
- Vorlagen: `data/Vorlagen/`
- Ausgaben: `data/LEKs/`
- Fachdokumente: `docs/`

## 7. Verteilungsüberblick

- Standardauslieferung: `release/LEK-Bastler_<Version>.zip`
- Inhalt: EXE + benötigte Ordner (`data/Aufgaben`, `data/Vorlagen`, `data/LEKs`, `docs`) + `README.md` + `LIZENZ.txt`

## 8. Qualitätssicherung (Kurzüberblick)

- Vor einem Release wird eine Regressionstest-Suite ausgeführt (`tools/test_regression_core.py`).
- Zusätzlich wird die Release-Checkliste aus `docs/RELEASE_QA_CHECKLISTE.md` durchlaufen.
- Für nicht-technische Prüfer steht zusätzlich `docs/ANWENDERPRUEFUNG_CHECKLISTE.md` zur Verfügung.
- Für kurze 5–10-Minuten-Tests steht ergänzend `docs/ANWENDERPRUEFUNG_KURZCHECKLISTE.md` zur Verfügung.

### Welche Checkliste ist die richtige?

- **Kurzcheckliste**: für einen schnellen Alltagstest mit kompaktem Gesamteindruck
- **Vollcheckliste**: für strukturierte Rückmeldungen zu Funktionalität, Verständlichkeit und Bedienbarkeit
