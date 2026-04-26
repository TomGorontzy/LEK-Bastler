# LEK-Bastler v2.2 - Export ins LEKs-Verzeichnis

## Neue Export-Funktionalität

Das LEK-Bastler Tool exportiert jetzt alle erstellten LEK-Dokumente automatisch ins `LEKs`-Unterverzeichnis.

## Änderungen

### Export-Pfad automatisiert

- **Automatische Verzeichniserstellung**: Das `LEKs`-Verzeichnis wird automatisch erstellt, falls es nicht existiert
- **Standard-Speicherort**: Export-Dialog öffnet sich direkt im `LEKs`-Verzeichnis
- **Konsistente Organisation**: Alle LEKs werden zentral an einem Ort gespeichert

### Verzeichnisstruktur

```text
LEK-Bastler/
├── LEKs/                    ← Alle exportierten LEK-Dokumente
│   ├── LIESMICH.md
│   └── [Exportierte LEK-Dateien]
├── Vorlagen/               ← LEK-Vorlagen
├── examples/               ← Beispiel-Aufgabendateien
└── [Tool-Dateien]

LEK-Bastler-Portable/
├── LEKs/                   ← Exportierte LEK-Dokumente (portable Version)
├── Vorlagen/              ← LEK-Vorlagen
├── dist/                  ← Ausführbare .exe-Dateien
│   └── LEK-Bastler-v2.2.exe
└── [Weitere Dateien]
```

### Dateinamen-Konvention

- **Mit LEK-Thema**: `LEK_[Thema]_[YYYYMM].docx`
- **Ohne LEK-Thema**: `LEK_Aufgaben_[YYYYMM].docx`

**Beispiele:**

```text
LEK_Auftragssteuerung_202511.docx
LEK_Koordination_202512.docx
LEK_Aufgaben_202511.docx
```

## Vorteile

### Organisation

- **Zentrale Sammlung**: Alle LEKs an einem definierten Ort
- **Keine verlorenen Dateien**: Benutzer können nicht versehentlich an falschen Orten speichern
- **Einfache Weitergabe**: Komplettes LEKs-Verzeichnis kann einfach kopiert/geteilt werden

### Benutzerfreundlichkeit

- **Automatisch vorausgewählter Pfad**: Export-Dialog startet im richtigen Verzeichnis
- **Konsistente Ablage**: Jeder Export landet im gleichen Ordner
- **Weniger Klicks**: Benutzer muss nicht jedes Mal zum LEKs-Ordner navigieren

### Portable Kompatibilität

- **Funktioniert in beiden Versionen**: Sowohl Python-Version als auch .exe-Version nutzen LEKs-Verzeichnis
- **Lokaler Pfad**: LEKs-Ordner wird neben der .exe bzw. im Python-Projektordner erstellt
- **Kein Systemzugriff**: Keine Administrator-Rechte erforderlich

## Implementation

### Code-Änderungen

- `main.py`: Automatische Erstellung des LEKs-Verzeichnisses vor Export
- `export_tasks`: Setzt `initialdir` auf den LEKs-Pfad
- Error-Handling für Verzeichniserstellung

### Rückwärtskompatibilität

- **Bestehende Workflows**: Benutzer können weiterhin in andere Verzeichnisse exportieren
- **Flexibilität**: Standard-Pfad ist LEKs, aber Änderung im Dialog möglich
- **Keine Breaking Changes**: Alle vorherigen Funktionen bleiben erhalten

### Zusätzliche Features

- **Automatisch vorausgewählter Pfad**: Export-Dialog startet im richtigen Verzeichnis
- **Intelligente Dateinamen**: Automatische Benennung basierend auf LEK-Thema und Datum
- **Überschreibschutz**: Benutzer kann Dateinamen bei Bedarf noch anpassen

### Systemkompatibilität

- **Funktioniert in beiden Versionen**: Sowohl Python-Version als auch .exe-Version nutzen LEKs-Verzeichnis
- **Relative Pfade**: Portable Version bleibt vollständig portabel
- **Konsistenz**: Gleiche Funktionalität in allen Umgebungen

## Implementation Details

### Technische Umsetzung

- `main.py`: Automatische Erstellung des LEKs-Verzeichnisses vor Export
- `initialdir=leks_dir`: Export-Dialog startet im LEKs-Verzeichnis
- `os.makedirs(leks_dir, exist_ok=True)`: Sichere Verzeichniserstellung

### Kompatibilität

- **Bestehende Workflows**: Benutzer können weiterhin in andere Verzeichnisse exportieren
- **Flexibilität**: Export-Dialog erlaubt Wechsel zu anderen Speicherorten
- **Keine Datenverluste**: Bestehende LEK-Dateien bleiben unberührt

## Versionen

- **LEK-Bastler-v2.2.exe**: Portable Version mit LEKs-Export-Funktionalität
- **Hauptverzeichnis**: Python-Version mit gleicher Funktionalität
- **Größe**: ~16,4 MB (portable .exe-Datei)

## Test-Workflow

1. Tool starten (Python-Version oder .exe-Datei)
2. Aufgabendatei laden
3. Aufgaben filtern und auswählen
4. "Exportieren" klicken
5. Export-Dialog öffnet sich im LEKs-Verzeichnis
6. Dateiname wird automatisch vorgeschlagen
7. LEK-Dokument wird im LEKs-Verzeichnis gespeichert
