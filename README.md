# LEK-Bastler-Portable v3.5.4

Ein professionelles Tool zum Auswählen und Exportieren von Aufgaben aus Word-Dokumenten mit robuster Strukturübertragung und vollständiger Formatierung.

## ✨ Hauptfunktionen

- **Intelligente Aufgaben-Extraktion**: Erkennt "Überschrift 1" formatierte Aufgaben automatisch
- **Vollständige Strukturerhaltung**: Tabellen, Listen, Formatierungen bleiben komplett erhalten
- **Präzise Metadaten-Extraktion**: Schwierigkeit und Keywords aus vollständigem Aufgabeninhalt
- **Optimierte GUI-Darstellung**: Spaltenbreiten angepasst für bessere Lesbarkeit
- **Vorlagen-System**: Automatische LEK-Erstellung mit professionellen Deckblättern
- **Portable Distribution**: Eigenständige .exe-Datei ohne Python-Installation
- **Seitennummerierung**: Automatische "Seite/Gesamtseiten" Nummerierung ab Seite 2
- **Benutzerfreundliche GUI**: Intuitive Bedienung mit korrekter Aufgabenvorschau und Metadaten

## 🚀 Installation

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

   ```bash
   python src/main.py
   ```

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

## 📝 Changelog

### v3.5 (2025) - Robuste Strukturübertragung

- ✅ **Verbesserte Kompatibilität**: Verwendet python-docx API statt direkte XML-Manipulation
- ✅ **Stabile Elementkopierung**: Robuste Übertragung von Paragraphen, Tabellen und Formatierungen
- ✅ **Fehlerbehandlung**: Graceful Fallbacks bei Kopierfehlern
- ✅ **Original-Strukturen**: Behält die komplette Aufgaben-Struktur aus dem Word-Dokument bei

### v3.4 (2025) - Komplette Strukturübertragung

- ✅ **Original-Überschriften**: Behält die ursprünglichen Aufgaben-Überschriften aus dem Word-Dokument bei
- ✅ **Vollständige Struktur**: Überträgt komplette Aufgabenbereiche von Überschrift 1 bis zur nächsten Überschrift 1
- ✅ **Korrekte Reihenfolge**: Überschriften und zugehörige Inhalte werden in der richtigen Reihenfolge übertragen
- ✅ **Formatierungserhalt**: Alle ursprünglichen Formatierungen bleiben erhalten

### v3.3 (2025) - Explizite Keyword-Extraktion

- ✅ **Explizite Keyword-Extraktion**: Erkennt "Schlüsselwörter:" Zeilen in Word-Dokumenten
- ✅ **Automatischer Fallback**: Wenn keine expliziten Schlüsselwörter gefunden werden
- ✅ **Schwierigkeitsgrad-Erkennung**: Unterstützt "Schwierigkeit:" Markierungen
- ✅ **Multi-Format Support**: Erkennt verschiedene Keyword-Formate (Deutsch/Englisch)

## 🛠️ Technische Details

- **Python**: 3.13+
- **GUI**: tkinter
- **Word-Verarbeitung**: python-docx mit XML-Manipulation
- **Build**: PyInstaller für portable .exe
- **Kompatibilität**: Windows 10/11

## 🔧 Build/Setup-Konvention

- Primäre Skripte: `src/build.ps1`, `src/setup.ps1`
- Root-Skripte `build.ps1` und `setup.ps1` sind Kompatibilitäts-Entrypoints.

## 📄 Lizenz

MIT License - Siehe src/LIZENZ.txt für Details
