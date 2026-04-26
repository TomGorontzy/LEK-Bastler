# LEK-Bastler Vorlagen-System

## Übersicht

Das erweiterte LEK-Bastler System unterstützt jetzt professionelle Word-Vorlagen für verschiedene Themenbereiche.

## Funktionsweise

### 1. Vorlagen-Ordner

- **Speicherort:** `Vorlagen/` (im LEK-Bastler Verzeichnis)
- **Format:** Word-Dokumente (.docx)
- **Benennungskonvention:** Thema im Dateinamen für automatische Zuordnung

### 2. Automatische Vorlagen-Auswahl

Das System erkennt automatisch passende Vorlagen basierend auf:

- LEK-Thema aus der Aufgabendatei (z.B. `Aufgaben_Auftragssteuerung-Koordination.docx`)
- Stichwörtern im Vorlagen-Dateinamen

### 3. Aufgaben-Einfügung

- **Position:** Ab der dritten Seite der Vorlage
- **Erhaltung:** Deckblatt und Seite 2 bleiben unverändert
- **Format:** Aufgaben werden strukturiert eingefügt

## Beispiel-Setup

### Vorlage erstellen

```text
1. Word-Dokument mit Deckblatt und Hinweisseiten erstellen
2. Als LEK-Vorlage_[Thema].docx speichern
3. In den Vorlagen-Ordner verschieben
```

### Aufgabendatei benennen

```text
Aufgaben_[Thema]-[Details].docx
Beispiel: Aufgaben_Auftragssteuerung-Koordination.docx
```

### Automatische Zuordnung

- System extrahiert "Auftragssteuerung" aus dem Dateinamen
- Sucht nach `LEK-Vorlage_Auftragssteuerung.docx`
- Verwendet diese Vorlage für den LEK-Export

## Vorlagen-Struktur

### Seite 1: Deckblatt

- Titel der LEK
- Firmenlogo/Schullogo
- Datum und Informationen

### Seite 2: Hinweise

- Bearbeitungsanweisungen
- Bewertungskriterien
- Zeitvorgaben

### Ab Seite 3: Aufgaben

- Automatisch eingefügte Aufgaben
- Originalformatierung erhalten
- Saubere Seitenumbrüche

## Fallback-Verhalten

Wenn keine passende Vorlage gefunden wird:

- Verwendet programmatisches Deckblatt (wie bisher)
- Alle Aufgaben werden direkt eingefügt
- Funktionalität bleibt vollständig erhalten

## Beispiel-Konfigurationen

### Aktuell verfügbar

- Auftragssteuerung und -koordination

### Erweiterbar für

- Mathematik/Naturwissenschaften
- Sprachen/Kommunikation
- Technik/IT
- Betriebswirtschaft
- Beliebige Fachbereiche

## Anpassung für andere Bereiche

Für Mathematik:

### Vorlagen

- `LEK-Vorlage_Mathematik.docx`
- Spezielle Formatierung für Formeln
- Angepasstes Layout für Berechnungen

### Aufgabendateien

- `Aufgaben_Mathematik-Grundlagen.docx`
- `Aufgaben_Mathematik-Analysis.docx`

## Vorteile

- **Professionelles Erscheinungsbild**
- **Zeitersparnis bei LEK-Erstellung**
- **Konsistente Corporate Identity**
- **Einfache Anpassung an verschiedene Themenbereiche**
- **Automatische Zuordnung ohne Benutzereingabe**

Das Vorlagen-System macht den LEK-Bastler noch professioneller und anpassbarer für verschiedene Bildungseinrichtungen und Unternehmen.
