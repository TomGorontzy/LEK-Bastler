# Dokumentation_Anwender

## 1. Zweck

Mit dem LEK-Bastler erstellen Sie aus vorbereiteten Aufgaben-Worddateien schnell eine Lernerfolgskontrolle (LEK).

## 2. Start

### Portable EXE

1. Öffnen Sie den ausgelieferten Ordner `LEK-Bastler-Portable_<Version>`.
2. Starten Sie `LEK-Bastler-Portable_<Version>.exe`.

### Python-Ausführung (Entwicklung)

1. Python-Umgebung vorbereiten
2. Anwendung starten über `src/main.py`

Hinweis zum aktuellen Stand: Version **3.5.0**.

## 3. Typischer Ablauf

1. **Aufgabensammlung laden** (`Aufgaben/*.docx`)
2. **Aufgaben filtern** (optional, z. B. über Suchbegriffe/Schwierigkeit)
3. **Aufgaben auswählen**
4. **Export starten**
5. Zieldatei im Ordner `LEKs/` speichern

Die vorgeschlagenen Dateinamen werden automatisch erzeugt und können bei Bedarf geändert werden.

## 4. Wichtige Hinweise

- Der Ordner `LEKs/` wird beim Export automatisch erstellt, falls er fehlt.
- Für beste Ergebnisse sollten Aufgaben sauber mit Überschriften strukturiert sein.
- Vorlagen liegen im Ordner `Vorlagen/`.
- Falls keine passende Spezialvorlage gefunden wird, wird die Standardvorlage verwendet.

## 5. Fehlerbehebung

### App startet nicht

- Prüfen, ob alle Dateien des Release-Pakets vorhanden sind.
- EXE einmal lokal auf Laufwerk statt direkt aus Cloud-Sync-Cache starten.
- Prüfen, ob Virenschutz/SmartScreen das Starten blockiert.

### Export schlägt fehl

- Sicherstellen, dass die Eingabedatei nicht in Word gesperrt ist.
- Schreibrechte im Zielordner prüfen.
- Prüfen, ob `Vorlagen/LEK-Vorlage.docx` vorhanden ist.
- Bei ungewöhnlichen Formatproblemen mit einer alternativen Aufgabenquelle gegenprüfen.

### Icon wird nicht korrekt angezeigt

- Neue EXE nach Build verwenden.
- Bei Bedarf Windows Explorer neu starten (Iconcache-Aktualisierung).

### LEKs-Ordner fehlt im ZIP

- In aktuellen Releases ist `LEKs/` enthalten.
- Beim Export wird `LEKs/` zusätzlich zur Laufzeit angelegt, falls nötig.

## 6. Datenablage

- Eingaben: `Aufgaben/`
- Vorlagen: `Vorlagen/`
- Ausgaben: `LEKs/`
- Fachdokumente: `docs/`

## 7. Verteilungsüberblick

- Standardauslieferung: `release/LEK-Bastler-Portable_<Version>.zip`
- Inhalt: EXE + benötigte Ordner (`Aufgaben`, `Vorlagen`, `LEKs`, `Dokumentation`) + `README.md` + `LIZENZ.txt`
