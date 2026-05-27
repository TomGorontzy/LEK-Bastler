# LEK-Bastler-Portable – Projektdokumentation

## 1. Ziel und Funktionsumfang

Der **LEK-Bastler-Portable** unterstützt die Erstellung von Lernerfolgskontrollen (LEKs) aus Word-Aufgabensammlungen.

Kernfunktionen:

- Aufgabenextraktion aus `.docx`
- Filterung/Selektion in GUI
- Export in LEK-Dokumente auf Basis von Vorlagen
- Erhaltung von Struktur und Formatierungen (inkl. Tabellen und Inhaltssteuerelemente)
- Automatische Ablage im Verzeichnis `data/LEKs/`

Aktueller Versionsstand: **3.6.1** (aus `src/build_version_info.txt`).

## 2. Projektstruktur (aktuell)

```text
LEK-Bastler-Portable/
├── src/
│   ├── main.py
│   ├── word_processor.py
│   ├── template_manager.py
│   ├── task_selector.py
│   └── app_icon.ico
│   ├── build.ps1
│   ├── LEK-Bastler-Portable.spec
│   ├── build_version_info.txt
│   ├── REQUIREMENTS.txt
│   └── LIZENZ.txt
├── data/
│   ├── Aufgaben/
│   ├── Vorlagen/
│   └── LEKs/
├── docs/                          # Aktuelle konsolidierte Dokumentation
│   ├── DOKUMENTATION_PROJEKT.md
│   ├── DOKUMENTATION_ANWENDER.md
│   └── DOKUMENTATION_TECHNIK.md
├── README.md
└── release/
```

## 3. Fachliche und technische Hauptmerkmale

### 3.1 Aufgaben-/Inhaltsübernahme

- Aufgaben werden strukturiert aus den Quelldokumenten entnommen.
- Erkennungsmodi sind konfigurierbar (`auto`, `headings`, `tables`, `mixed`).
- Inhalte werden in der Exportphase inklusive komplexer Elemente übertragen.
- Metadaten wie Schwierigkeit/Keywords werden für Auswahl- und Exportzwecke verwendet.

### 3.5 Regelwerk und Parser (Sprints 2–3)

- Fachregeln sind zentral in `data/config/import_rules.json` gebündelt.
- Feldaliase, Pflichtfelder und Blockierregeln werden konsistent in GUI und Verarbeitung genutzt.
- Parser-Pipeline trennt Strukturerkennung, Moduswahl/Fallback und Normalisierung.

### 3.2 Vorlagen-System

- Vorlagen liegen im Ordner `data/Vorlagen/`.
- LEK-Themen können für die Vorlagenauswahl und Benennung genutzt werden.
- Platzhalter in Vorlagen werden beim Export ersetzt.

### 3.3 Export in `data/LEKs/`

- Der LEK-Export nutzt standardmäßig das Verzeichnis `data/LEKs/`.
- Das Verzeichnis wird bei Bedarf zur Laufzeit automatisch erzeugt.

### 3.4 Icon-Integration

- EXE-Icon wird über Spec eingebettet (`src/app_icon.ico`).
- GUI-Fenstericon wird zur Laufzeit gesetzt.
- Die ICO enthält mehrere Auflösungen für stabile Windows/Explorer-Darstellung.

## 4. Build- und Release-Prozess

Build-Skript: `src/build.ps1`

### 4.1 Artefakte

- PyInstaller-EXE (Zwischenartefakt): `dist/LEK-Bastler-Portable.exe`
- Deploy-Ordner: `dist/LEK-Bastler-Portable_<Version>/`
- Deploy-EXE: `LEK-Bastler-Portable_<Version>.exe`
- Release-ZIP: `release/LEK-Bastler-Portable_<Version>.zip`

Beispiel für aktuelle Version:

- `dist/LEK-Bastler-Portable_3.5.6/LEK-Bastler-Portable_3.5.6.exe`
- `release/LEK-Bastler-Portable_3.5.6.zip`

### 4.2 Deploy-Inhalt

- EXE
- `data/Aufgaben/`
- `docs/`
- `data/Vorlagen/`
- `data/LEKs/` (inkl. Platzhalterdatei)
- `README.md`
- `LIZENZ.txt` (aus `src/LIZENZ.txt`)

## 5. Architekturüberblick

- `src/main.py`: GUI, Dateiauswahl, Filter, Exportfluss
- `src/task_selector.py`: Selektions-/Filterlogik
- `src/word_processor.py`: Word-Einlesen, Extraktion, Dokumenterstellung
- `src/template_manager.py`: Vorlagenauswahl, Platzhalterersetzung, Einfügelogik

## 6. Qualitätssicherung und Regression

Empfohlene Mindestprüfungen:

1. Start aus Python (`src/main.py`)
2. Build via `src/build.ps1` erfolgreich
3. Deploy-Struktur vollständig
4. ZIP-Inhalt vollständig
5. Testexport mit einer Datei aus `data/Aufgaben/`
6. Icon-Prüfung (EXE + Fenstertitel)

Automatisierte Regression:

- `tests/test_regression_core.py` deckt zentrale Kernpfade (R1–R6) ab.
- Testfallmatrix: `memos/MEMO_REGRESSIONSTEST_MATRIX.md`.
- Release-QA-Checkliste: `docs/RELEASE_QA_CHECKLISTE.md`.

## 7. Historischer Kontext

Historische Inhalte wurden in diese Dokumentation fachlich übernommen und auf den aktuellen Stand (Version, Struktur, Namen) gebracht.
