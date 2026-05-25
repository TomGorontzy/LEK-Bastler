# DOKUMENTATION_TECHNIK

## 1. Komponenten

- GUI: `src/main.py` (tkinter/ttk)
- Extraktion/Export: `src/word_processor.py`
- Vorlagenlogik: `src/template_manager.py`
- Selektion: `src/task_selector.py`
- Anwendungssymbol: `src/app_icon.ico`

Aktueller Versionsstand: **3.5.0**.

## 2. Build-Konfiguration

- Spec: `src/LEK-Bastler-Portable.spec`
  - Entry: `src/main.py`
  - Pathex: `src`
  - Icon: `src/app_icon.ico`
  - Datas: `data`, `src/app_icon.ico`
  - EXE-Basisname: `LEK-Bastler-Portable`

- Build-Skript: `src/build.ps1`
  - Führt PyInstaller aus
  - Erzeugt versionierten Deploy-Ordner `dist/LEK-Bastler-Portable_<Version>`
  - Erzeugt versionierte EXE `LEK-Bastler-Portable_<Version>.exe`
  - Erzeugt versioniertes ZIP in `release/LEK-Bastler-Portable_<Version>.zip`
  - Nutzt .NET-ZIP-Erstellung (`ZipFile.CreateFromDirectory`) für zuverlässige Inhalte

## 3. Pfadkonzept

- Entwicklungsmodus: Pfade relativ zu `__file__` in `src/`
- EXE-Modus: Pfade relativ zu `sys.executable` (Deploy-Ordner)
- Icon-Fallback bei PyInstaller: `sys._MEIPASS`
- Exportziel: Verzeichnis `data/LEKs/` (wird bei Bedarf automatisch angelegt)

## 4. Inhaltserhaltung beim Export

- Übernahme der Aufgabenstrukturen inklusive Tabellen
- Erhaltung von Formatierungsinformationen
- Übernahme von Inhaltssteuerelementen (SDT)
- Kontextabgleich zwischen Quelle und Zielvorlage (Styles/Nummerierung)

## 5. Konventionen

- Source-Dateien liegen unter `src/`
- Laufzeitdaten (`data/Aufgaben`, `data/Vorlagen`, `data/LEKs`) liegen unter `data/`
- Build-Artefakte liegen unter `dist/` und `release/`
- Anwenderdoku und technische Doku liegen unter `docs/`

## 6. Verteilungsoptionen

Aus den historischen Dokumenten übernommen und auf aktuellen Stand angepasst:

- Portable Release als ZIP (primärer Weg)
- Lokales Build für interne Tests
- Netzwerkbereitstellung des entpackten Release-Ordners

## 7. Regression-Checkliste

1. `src/build.ps1` erfolgreich
2. `dist/LEK-Bastler-Portable_<Version>/` vollständig
3. `release/LEK-Bastler-Portable_<Version>.zip` vollständig
4. EXE-Icon korrekt
5. GUI-Fenstericon korrekt
6. Exportfunktion erstellt Word-Datei unter `data/LEKs/`
7. Vorlagenersetzung und Aufgabenübernahme fachlich plausibel
