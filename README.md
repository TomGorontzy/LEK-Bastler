# LEK-Bastler-Portable v3.5.0

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
2. Abhängigkeiten installieren:

   ```bash
   pip install python-docx lxml
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
├── Aufgaben/              # Quell-Word-Dokumente mit Aufgabensammlungen
├── Vorlagen/              # LEK-Vorlagen für verschiedene Themen
├── LEKs/                  # Exportierte Lernerfolgskontroll-Dokumente
└── docs/                  # Aktuelle Projektdokumentation
```

## 📝 Changelog

### v3.2 (29.11.2025) - GUI-Optimierungen

- **🎨 Optimierte Spaltenbreiten**: Nr (50px), Titel (300px), Schwierigkeit (100px), Suchbegriffe (250px)
- **🔧 Verbesserte Keywords-Anzeige**: Saubere Darstellung der Schlüsselwörter in der Aufgabenliste
- **🔧 Robustere GUI**: Bessere Fehlerbehandlung bei leeren Keyword-Listen

### v3.0 (29.11.2025) - Metadaten-Fix

- **🔧 KRITISCHER FIX**: Schwierigkeitsgrad-Anzeige korrigiert - basiert nun auf vollständigem Aufgabeninhalt
- **🔧 KRITISCHER FIX**: Schlüsselwörter-Extraktion korrigiert - erfasst nun alle relevanten Begriffe aus gesamter Aufgabe  
- **🔧 Optimierte Metadaten-Berechnung**: Schwierigkeit und Keywords werden nach vollständiger Inhalterfassung berechnet
- **🔧 Entfernung redundanter Verarbeitung**: Sauberere und effizientere Extraktion

### v2.9 (29.11.2025)

- **Verbesserte Aufgaben-Extraktion**: Erfasst ALLE Inhalte zwischen "Überschrift 1" Formatierungen
- **Entfernung "Aufgabe x:" Präfix**: Verwendet nur noch ursprüngliche Titel
- **Seitennummerierung**: Format "Seite/Gesamtseiten" rechtsbündig ab Seite 2
- **Vollständige Strukturerhaltung**: XML-basierte Element-Übertragung

### v2.8 (29.11.2025)

- **Prioritäre Überschrift-Erkennung**: "Überschrift 1" Format hat Vorrang
- **Robuste Element-Verarbeitung**: Auch leere Absätze für korrekte Formatierung
- **Verbesserte Gruppierung**: Komplette Abschnitte zwischen Überschriften

### v2.7 (28.11.2025)

- **Pfad-Korrekturen**: Fixes für .exe vs Python-Ausführung
- **GUI-Verbesserungen**: Korrekte Verzeichnis-Navigation
- **Portable Distribution**: Cleanere Verzeichnisstruktur

## 📋 Changelog

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

### v3.2 (2024) - GUI-Optimierung  

- ✅ **Spaltenbreiten optimiert**: Nr (50px), Titel (300px), Schwierigkeit (100px), Suchbegriffe (250px)
- ✅ **Verbesserte Darstellung**: Bessere Lesbarkeit der Aufgaben-Metadaten

### v3.0 (2024) - Metadaten-Korrektur

- ✅ **Korrigierte Metadaten-Anzeige**: Zeigt tatsächlich extrahierte Keywords und Schwierigkeitsgrade an
- ✅ **Verbesserte Extraktion**: Robustere Erkennung von Aufgaben-Eigenschaften

## 🛠️ Technische Details

- **Python**: 3.13+
- **GUI**: tkinter
- **Word-Verarbeitung**: python-docx mit XML-Manipulation
- **Build**: PyInstaller für portable .exe
- **Kompatibilität**: Windows 10/11

## 📄 Lizenz

MIT License - Siehe LIZENZ.txt für Details
