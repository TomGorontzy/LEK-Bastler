# Idee: Intuitive und fehlerarme Aufgabenübernahme in LEK-Bastler-Portable

## Ausgangslage

Die aktuelle Erstellung und Übernahme von Aufgaben ist funktional, aber für unerfahrene Anwender noch nicht ausreichend intuitiv.

Zusätzlich enthalten Quell-Dokumente teilweise einleitende Texte vor den eigentlichen Aufgaben. Diese Struktur führt aktuell leicht zu Fehlinterpretationen beim Import.

## Ziel

Die Aufgabenübernahme soll so geführt werden, dass auch unerfahrene Anwender reproduzierbar das gewünschte Ergebnis in den LEKs erhalten.

## Vorschlag

1. **Klare Datenstruktur je Aufgabe**
   - Intro/Einleitung (optional)
   - Aufgabenstellung (Pflicht)
   - Lösungsmöglichkeit/Hinweis (optional)
   - Schlagworte
   - Schwierigkeitsgrad (leicht/mittel/schwer)

2. **Geführter Import-Workflow (Wizard)**
   - Schritt 1: Quelle wählen
   - Schritt 2: Erkannte Aufgaben prüfen
   - Schritt 3: Vorschau der LEK-Ausgabe
   - Schritt 4: Übernahme bestätigen

3. **Sicherheitsleitplanken**
   - Warnung bei 0 erkannten Aufgaben
   - Prüfung leerer Aufgabenblöcke
   - Klare, handlungsorientierte Fehlermeldungen

4. **Word-Gerüst als Eingabehilfe**
   - Vorlage vorhanden: `data/Vorlagen/AUFGABEN_GERUEST_WORD.docx`
   - Ziel: standardisierte Erfassung bei gleichzeitig hoher Flexibilität (Grafiken, Tabellen, Formatierung)

## Erste Akzeptanzkriterien

- Intro-Texte werden nicht versehentlich als Aufgaben importiert.
- Nutzer sieht vor der finalen Übernahme eine verlässliche Vorschau.
- Bei typischen Fehlern wird verständlich auf die Ursache hingewiesen.
- Ergebnis in LEK entspricht der bestätigten Vorschau.

## Mitmachen

Beiträge sind willkommen, z. B. zu:

- Parser-Regeln (Intro vs. Aufgabe)
- UX-/Dialogfluss
- Validierung und Fehlermeldungen
- Beispielquellen und Regressionstests

Gern direkt Vorschläge, Mockups oder PRs verlinken.

## Update 2026-05-26 – Schwierigkeitsgrad-Regel geschärft

- Konfigurierbare Difficulty-Regeln ergänzt (`difficulty_rules` in `data/config/import_rules.json`):
   `allowed_values`, `aliases`, `block_export_on_inconsistent`.
- Parser normalisiert Difficulty-Werte regelbasiert (inkl. Alias-Mapping).
- Inkonsistente Mehrfachwerte (z. B. `leicht | mittel | schwer`) werden robust erkannt.
- GUI nutzt dieselben Regelwerte bei der Metadaten-Eingabe.
- Bulk-Import besitzt jetzt eine Korrekturschleife bei abgebrochener Metadaten-Eingabe:
   erneut bearbeiten / Datei überspringen / Serie stoppen.
- Export blockiert optional bei inkonsistenter oder ungültiger Schwierigkeit.

## Update 2026-05-26 – Titel als Pflichtfeld aufgenommen

- Ja, das Feld `Titel` wurde als Template-Pflichtfeld aufgenommen.
- Neue Regelkonfiguration `template_rules.required_fields` enthält jetzt standardmäßig:
   `id`, `aufgabenstellungpflicht`, `titel`.
- Fehlt ein Pflichtfeld (z. B. `Titel`), wird eine Diagnosewarnung erzeugt (`Pflichtfeld fehlt: ...`).
- Bei aktivem `template_rules.block_export_on_missing_required=true` wird der Export blockiert,
   bis die Pflichtfelder in der Quelle ergänzt und neu geladen wurden.

## Update 2026-05-26 – Migrationshilfe für Altdokumente

- Beim Laden einer Aufgabensammlung erkennt die GUI fehlende Titel (`Pflichtfeld fehlt: Titel`).
- Anwender erhalten eine direkte Rückfrage zur automatischen Korrektur.
- Bei Zustimmung werden fehlende Titel aus der Aufgabenstellung abgeleitet,
  die Sammlung gespeichert und automatisch neu geladen.
- Vor der Änderung wird eine zeitgestempelte Backup-Datei erstellt.
- Der Erfolgsdialog zeigt zusätzlich die betroffenen Aufgaben-IDs
   (bei vielen Einträgen mit gekürzter Vorschau).
- Zusätzlich wird eine kompakte Vorschau `ID -> abgeleiteter Titel` angezeigt
   (erste Einträge, bei Bedarf gekürzt).

## Update 2026-05-26 – Anzeige von Haupt-/Nebennummern

- Die Spalte **Nr** zeigt jetzt die erkannte Aufgaben-Nummer aus der Quelle,
   inklusive Nebennummern (z. B. `1`, `1.1`, `1.2`).
- Nummern werden aus Überschriften (`Aufgabe 1.1 ...`) bzw. aus der Tabellen-ID
   (`ID`) extrahiert.
- Intern bleibt eine stabile numerische technische ID für Auswahl/Freigabe erhalten,
   sodass die Wizard-Logik robust bleibt.
- Zusätzliche Regel: Wird eine Aufgabe als Einleitung/Kontext erkannt,
  wird die sichtbare Nummer auf Nebennummer `.0` normalisiert
  (z. B. `1` → `1.0`). Bereits vorhandene Nebennummern wie `1.1` bleiben unverändert.

## Update 2026-05-26 – Export strukturierter Aufgaben als Fließtext

- Beim LEK-Export werden strukturierte Aufgaben-Tabellen nicht mehr als Tabelle übernommen.
- Stattdessen wird ausschließlich der Inhalt der rechten Spalte als Fließtext ausgegeben
   (inkl. Textformatierungen pro Absatz/Run).
- Reihenfolge der Blöcke:
   1. Titel
   2. (Leerzeile)
   3. Intro/Einleitung (optional)
   4. (Leerzeile)
   5. Aufgabenstellung
   6. (Leerzeile)
   7. Hinweis (optional)
   8. (Leerzeile)
   9. Punkte (optional)

### Nachfix

- Feldalias-Erkennung im Parser/Export erweitert (`introeinleitung`,
   `loesungsmoeglichkeithinweis`, weitere Varianten ohne "optional").
- Dadurch werden Introtexte bei Aufgaben mit Nebennummer `.0`
   zuverlässig in die LEK übernommen.

## Update 2026-05-26 – Schritt-3-Gesamtvorschau exportnah ergänzt

- Wizard-Schritt 3 enthält jetzt eine **LEK-Gesamtvorschau** über alle freigegebenen Aufgaben
   (Button: `LEK-Gesamtvorschau`).
- Die Vorschau nutzt dieselbe Blockreihenfolge wie der Export strukturierter Aufgaben:
   1. Titel
   2. Intro/Einleitung (optional)
   3. Aufgabenstellung
   4. Hinweis (optional)
   5. Punkte (optional)
- Für strukturierte Aufgaben wurde ein **Delta-Check** ergänzt:
   fehlende optionale Blöcke werden explizit angezeigt
   (`Delta-Check (optional nicht vorhanden): ...`).
- Ziel: bessere Nachvollziehbarkeit vor dem finalen Export
   und engeres Verhalten „Vorschau = Export“.
