# LEK-Bastler-Portable

Ein Desktop-Werkzeug zum Auswählen und Exportieren von Aufgaben aus Word-Dokumenten mit robuster Strukturübertragung, Vorlagenlogik und portabler Windows-Auslieferung.

Der fachliche Schwerpunkt liegt auf der schnellen Erstellung von Lernerfolgskontrollen (LEKs) aus vorhandenen Aufgabensammlungen.

Der aktuelle Versionsstand wird über `src/build_version_info.txt` und die GitHub-Releases gepflegt.

## ✨ Hauptfunktionen

- **Intelligente Aufgaben-Extraktion**: Erkennt "Überschrift 1" formatierte Aufgaben automatisch
- **Vollständige Strukturerhaltung**: Tabellen, Listen, Formatierungen bleiben komplett erhalten
- **Präzise Metadaten-Extraktion**: Schwierigkeit und Keywords aus vollständigem Aufgabeninhalt
- **Optimierte GUI-Darstellung**: Spaltenbreiten angepasst für bessere Lesbarkeit
- **Vorlagen-System**: Automatische LEK-Erstellung mit professionellen Deckblättern
- **Portable Distribution**: Eigenständige .exe-Datei ohne Python-Installation
- **Seitennummerierung**: Automatische "Seite/Gesamtseiten" Nummerierung ab Seite 2
- **Benutzerfreundliche GUI**: Intuitive Bedienung mit korrekter Aufgabenvorschau und Metadaten

## 🚀 Schnellstart

### Option 1: Portable Version (Empfohlen)

1. Laden Sie `LEK-Bastler-Portable_<Version>.zip` herunter
2. Entpacken Sie das Archiv
3. Führen Sie `LEK-Bastler-Portable_<Version>.exe` aus

### Option 2: Python-Version

1. Python 3.13+ installieren
2. Virtuelle Umgebung und Abhängigkeiten installieren:

   ```powershell
   .\src\setup.ps1
   ```

3. Programm starten:

   ```powershell
   python src/main.py
   ```

## 🎯 Typischer Ablauf

1. Aufgabensammlung laden
2. Aufgaben sichten und filtern
3. Geeignete Aufgaben auswählen
4. Optional ein LEK-Thema vergeben
5. Export in `data/LEKs/` starten

![Anwenderablauf LEK-Bastler-Portable](docs/diagramme/anwender_ablauf.svg)

## 📋 Verwendung

1. **Word-Datei laden**: Klicken Sie "Durchsuchen" und wählen Sie Ihre Aufgabensammlung
2. **Aufgaben auswählen**: Wählen Sie gewünschte Aufgaben aus der Liste
3. **LEK-Thema eingeben**: Optional für thematische Gruppierung
4. **Exportieren**: Klicken Sie "LEK erstellen" für automatischen Export

## 🗂️ Projektstruktur

```text
LEK-Bastler-Portable/
├── src/
│   ├── main.py             # Hauptprogramm mit GUI
│   ├── word_processor.py   # Word-Dokumentverarbeitung mit XML-Zugriff
│   ├── template_manager.py # Vorlagen-Management und LEK-Erstellung
│   └── task_selector.py    # Filterlogik
│   ├── build.ps1
│   ├── setup.ps1
│   ├── LEK-Bastler-Portable.spec
│   ├── build_version_info.txt
│   ├── REQUIREMENTS.txt
│   └── LIZENZ.txt
├── data/
│   ├── Aufgaben/           # Quell-Word-Dokumente mit Aufgabensammlungen
│   ├── Vorlagen/           # LEK-Vorlagen für verschiedene Themen
│   └── LEKs/               # Exportierte Lernerfolgskontroll-Dokumente
└── docs/                  # Aktuelle Projektdokumentation
```

## 📝 Änderungsverlauf

Der konsolidierte Änderungsverlauf wird zentral in `docs/CHANGELOG.md` gepflegt.

## 🛠️ Technische Details

- **Python**: 3.13+
- **GUI**: tkinter
- **Word-Verarbeitung**: python-docx mit XML-Manipulation
- **Build**: PyInstaller für portable .exe
- **Kompatibilität**: Windows 10/11

## 🔧 Build/Setup-Konvention

- Primäre Skripte: `src/build.ps1`, `src/setup.ps1`
- Root-Skripte `build.ps1` und `setup.ps1` sind Kompatibilitäts-Entrypoints.

## 🚀 GitHub Release-Automation

- Workflow: `.github/workflows/release.yml`
- Trigger: Push eines Tags im Format `v*` (z. B. `v3.6.1`)
- Guard: Tag-Version muss exakt zur `FileVersion` aus `src/build_version_info.txt` passen
- Ergebnis: Build via `src/setup.ps1` + `src/build.ps1`, Upload von EXE/ZIP/Release Notes als Assets, Release-Titel-Schema `LEK-Bastler-Portable v...`

Für jeden Build/Release wird zusätzlich eine Notes-Datei erzeugt:

- `release/RELEASE_NOTES_v<version>.md`

Sie dient als Basistext für den GitHub-Release-Eintrag.

## ✅ Qualitätssicherung

- Automatisierte Regressionstests: `tests/test_regression_core.py`
- Testfallmatrix: `memos/MEMO_REGRESSIONSTEST_MATRIX.md`
- Release-Checkliste: `docs/RELEASE_QA_CHECKLISTE.md`
- Release-Prozess: `docs/DOKUMENTATION_RELEASES.md`

## 📚 Dokumentation

- Projekt: `docs/DOKUMENTATION_PROJEKT.md`
- Anwender: `docs/DOKUMENTATION_ANWENDER.md`
- Diagramme: `docs/DOKUMENTATION_DIAGRAMME.md`
- Technik: `docs/DOKUMENTATION_TECHNIK.md`
- Releases: `docs/DOKUMENTATION_RELEASES.md`
- Änderungen: `docs/CHANGELOG.md`
- Release-QA: `docs/RELEASE_QA_CHECKLISTE.md`
- Smoke-Test-Protokoll: `docs/RELEASE_SMOKETEST_PROTOKOLL.md`

Der empfohlene Einstieg ist: `README.md` → `docs/DOKUMENTATION_ANWENDER.md` → `docs/DOKUMENTATION_RELEASES.md`.

Die Artefakte zu Qualitätssicherung, Release-Ablauf und Änderungsverlauf bilden zusammen den Mindest-Qualitätsrahmen vor jedem Release.

## 📄 Lizenz

MIT License - Siehe src/LIZENZ.txt für Details
