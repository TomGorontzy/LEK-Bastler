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

Hinweis zum aktuellen Stand: Version **3.7.7**.

## 3. Typischer Ablauf

1. **Aufgabensammlung laden** (`data/Aufgaben/*.docx`)
2. **Aufgaben filtern** (optional, z. B. über Suchbegriffe/Schwierigkeit)
3. **Aufgaben auswählen und freigeben**
4. **Optional Gesamtausgabe prüfen**
5. **Export starten**
6. Zieldatei im Ordner `data/LEKs/` speichern
7. Automatisch erzeugte Protokolldatei (`.txt`) neben der LEK prüfen (enthält verwendete Aufgabennummern)

Die vorgeschlagenen Dateinamen werden automatisch erzeugt und können bei Bedarf geändert werden.

## 3a. Visualisierter Ablauf

![Anwenderablauf LEK-Bastler](diagramme/anwender_ablauf.svg)

Das Diagramm zeigt den typischen Pfad von der Aufgabensammlung bis zum Export in `data/LEKs/`.
Es bildet zusätzlich die Freigabe-Gating-Logik ab ("Weiter gesperrt", solange keine Freigaben vorliegen).

## 4. Wichtige Hinweise

- Der Ordner `data/LEKs/` wird beim Export automatisch erstellt, falls er fehlt.
- Zu jeder exportierten LEK wird automatisch eine Protokolldatei mit gleichem Dateinamen und Endung `.txt` erzeugt.
- Die Protokolldatei enthält eine Liste der tatsächlich verwendeten Aufgabennummern und liegt im gleichen Ordner wie die exportierte LEK.
- Für beste Ergebnisse sollten Aufgaben sauber mit Überschriften oder als strukturierte 2-Spalten-Tabellen aufgebaut sein.
- Zusammengehörige Aufgabenfamilien werden automatisch gemeinsam behandelt: Wenn eine Hauptaufgabe mit Unteraufgaben vorliegt (z. B. `1.0`, `1.1`, `1.2`), übernimmt der LEK-Bastler diese bei Auswahl, Freigabe und Export immer als vollständige Gruppe.
- Wenn Sie nur eine Unteraufgabe markieren, erweitert die Anwendung die Auswahl automatisch auf die restlichen Aufgaben derselben Hauptnummer.
- Vorlagen liegen im Ordner `data/Vorlagen/`.
- Relevante Vorlagen im aktuellen Stand:
  - `data/Vorlagen/LEK-Vorlage.docx` (LEK-Export)
  - `data/Vorlagen/AUFGABEN-Vorlage.docx` (Muster für Aufgabensammlungen)
- Falls keine passende Spezialvorlage gefunden wird, wird die Standardvorlage verwendet.
- Vorschau und Export nutzen dieselbe fachliche Reihenfolge bei strukturierten Aufgaben (Titel → Intro → Aufgabenstellung → Hinweis → Punkte).
- Generierte LEKs sollten vor einer Veröffentlichung kurz auf ungünstige Seitenumbrüche geprüft und bei Bedarf manuell in Word nachbearbeitet werden.

## 4a. Regelwerk im Alltag (ab v3.6.0)

- Viele Import-/Validierungsregeln sind zentral in `data/config/import_rules.json` konfigurierbar.
- Änderungen an Aliasen oder Pflichtfeldern können in der Regel ohne Codeänderung erfolgen.
- Bei Regelverstößen (z. B. fehlende Kategorie, inkonsistente Schwierigkeit) kann der Export bewusst blockiert werden.

## 4b. Externe Großtabellen referenzieren

- Für sehr große Tabellen oder tabellarische Anhänge kann in einer Aufgabe ein Marker wie `<<tabelle=Kalkulationsschema>>` oder `<<tabelle=Kalkulationsschema.docx>>` verwendet werden.
- Die referenzierte Datei wird automatisch im lernbereichsspezifischen Unterordner der geladenen Aufgabensammlung gesucht, z. B. bei `Aufgaben_Auftragssteuerung und -koordination.docx` unter `data/Aufgaben/Auftragssteuerung und -koordination/`.
- Der Marker selbst wird nicht in die LEK übernommen; stattdessen wird das referenzierte Dokument eingebettet.
- Die Anwendung wählt automatisch Hoch- oder Querformat. Optional kann zusätzlich `<<tabelle_format=landscape>>`, `<<tabelle_format=portrait>>` oder `<<tabelle_format=auto>>` gesetzt werden.
- Extern referenzierte Inhalte werden nach Möglichkeit jeweils auf einer eigenen Seite in die LEK übernommen.
- Wenn die referenzierte Datei fehlt, wird der Export mit einer klaren Meldung blockiert.

## 5. Fehlerbehebung

### App startet nicht

- Prüfen, ob alle Dateien des Windows-Releases vorhanden sind.
- EXE einmal lokal auf Laufwerk statt direkt aus Cloud-Sync-Cache starten.
- Prüfen, ob Virenschutz/SmartScreen das Starten blockiert.

### Export schlägt fehl

- Sicherstellen, dass die Eingabedatei nicht in Word gesperrt ist.
- Schreibrechte im Zielordner prüfen.
- Prüfen, ob `data/Vorlagen/LEK-Vorlage.docx` vorhanden ist.
- Prüfen, ob zusammengehörige Teilaufgaben fachlich vollständig gepflegt sind; da Aufgabenfamilien gemeinsam exportiert werden, können Warnungen in einer Unteraufgabe die gesamte Gruppe betreffen.
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
