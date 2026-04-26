# LEKs Export-Verzeichnis

Dieses Verzeichnis wird für die exportierten Lernerfolgskontroll-Dokumente (LEKs) verwendet.

## 🔧 Funktionsweise

- Das LEK-Bastler Tool exportiert alle erstellten LEK-Dokumente automatisch in dieses Verzeichnis
- Der Export-Dialog öffnet sich standardmäßig hier
- Dateien werden automatisch mit Thema und Datum benannt (z.B. "LEK_Auftragssteuerung_202511.docx")
- **Seitennummerierung**: Automatisch "Seite/Gesamtseiten" ab Seite 2 (rechtsbündig)
- **Formaterhaltung**: Vollständige Übernahme aller Strukturen aus Quell-Dokumenten

## ✅ Vorteile

- **Zentrale Organisation**: Alle LEKs an einem Ort
- **Automatische Sortierung**: Durch Dateinamen-Konvention leicht auffindbar
- **Keine verlorenen Dateien**: Klare Struktur verhindert versehentliches Speichern an falschen Orten
- **Portable Nutzung**: Funktioniert sowohl in der Python-Version als auch in der portablen .exe-Datei
- **Professionelle Ausgabe**: Mit Deckblatt, Seitennummern und korrekter Formatierung

## 📋 Dateinamen-Konvention

- **Mit LEK-Thema**: `LEK_[Thema]_[YYYYMMDD]_[HHMM].docx`
- **Ohne LEK-Thema**: `LEK_Aufgaben_[YYYYMMDD]_[HHMM].docx`

### Beispiele

- `LEK_Auftragssteuerung_20251129_1430.docx`
- `LEK_Koordination_20251215_0900.docx`
- `LEK_Aufgaben_20251129_1045.docx`

## 🆕 Features v2.9

- **Saubere Überschriften**: Keine "Aufgabe x:" Präfixe mehr
- **Vollständige Inhalte**: Alle Elemente zwischen Überschriften werden erfasst
- **Automatische Seitennummerierung**: "1/5", "2/5" etc. ab Seite 2
- **Strukturtreue**: Tabellen, Listen, Formatierungen bleiben erhalten

