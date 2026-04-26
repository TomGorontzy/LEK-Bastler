# LEK-Bastler v2.1 - Formatierungserhalt

## Neue Funktionen

### Vollständiger Formatierungserhalt

Das LEK-Bastler Tool v2.1 unterstützt jetzt die vollständige Erhaltung der Word-Formatierung beim Kopieren von Aufgaben.

**Was wird erhalten:**

- **Paragraph-Stile**: Original-Formatvorlagen werden übernommen
- **Text-Ausrichtung**: Links-, rechts-, zentrierte oder Blocksatz-Ausrichtung
- **Einzüge**: Linke, rechte und Erstzeileneinzüge
- **Abstände**: Abstände vor und nach Absätzen
- **Schriftarten**: Name, Größe und Farbe der Schriftarten
- **Text-Formatierung**: Fett, kursiv, unterstrichen
- **Erweiterte Formatierung**: Hochgestellt, tiefgestellt, durchgestrichen, Kapitälchen

### Technische Implementierung

#### Aufgaben-Datenstruktur erweitert

- `original_paragraphs`: Speichert die ursprünglichen Word-Paragraph-Objekte
- Fallback auf reine Text-Darstellung, falls keine Original-Paragraphen verfügbar

#### Formatierungs-Kopie-Methode

```python
def _copy_paragraphs_with_formatting(self, target_doc, source_paragraphs):
```

Diese Methode kopiert Word-Paragraphen mit vollständiger Formatierung von einem Dokument ins andere.

### Anwendung

1. **Aufgaben-Erkennung**: Basierend auf "Überschrift 1" Formatvorlage
2. **Formatierungserhalt**: Beim Export werden Original-Formatierungen übernommen
3. **Template-Integration**: Formatierte Aufgaben werden professionell in LEK-Vorlagen eingesetzt

### Kompatibilität

- **Aufwärtskompatibel**: Alte Aufgaben ohne Formatierung werden weiterhin unterstützt
- **Template-System**: Vollständig kompatibel mit dem bestehenden Vorlagen-System
- **Portable Version**: Alle Verbesserungen sind in der ausführbaren .exe-Datei enthalten

### Test-Szenarien

Um die Formatierung zu testen:

1. Erstelle eine Word-Datei mit verschiedenen formatierten Aufgaben
2. Verwende "Überschrift 1" für Aufgabentitel
3. Formatiere den Aufgabeninhalt mit verschiedenen Schriftarten, Farben, etc.
4. Exportiere mit dem LEK-Bastler Tool
5. Überprüfe, dass alle Formatierungen erhalten bleiben

### Dateien

- **Hauptversion**: `/LEK-Bastler/` (Python-Umgebung erforderlich)
- **Portable Version**: `/LEK-Bastler-Portable/dist/LEK-Bastler-v2.1.exe` (keine Installation erforderlich)
