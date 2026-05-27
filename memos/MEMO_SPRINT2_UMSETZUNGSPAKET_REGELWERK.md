# MEMO: Sprint-2 Umsetzungspaket – Regelwerk und Konfiguration

## Ziel

Sprint 2 überführt die inzwischen gewachsenen Import-, Validierungs- und Exportregeln in ein klar strukturiertes, zentral pflegbares Regelwerk. Fachliche Feinjustierung soll möglichst ohne Codeänderungen möglich sein.

## Scope (Sprint 2)

- Ausbau von `data/config/import_rules.json`
- Zentrale Konfiguration für Pflichtfelder, Blockierregeln und Feldalias-Mappings
- Vereinheitlichung der Regelabfragen in GUI und Verarbeitung
- Dokumentierte Governance für neue Sammlungen

## Nicht-Scope (ab Sprint 3)

- Tiefer Parser-Refactor für gemischte Dokumentstrukturen
- Große Regressionstest-Suite mit Fixture-Katalog
- Visueller GUI-Feinschliff jenseits regelbezogener Hinweise

## Bestehende Basis aus Sprint 1

- `src/import_wizard.py` etabliert den Sitzungs- und Freigabefluss.
- `src/word_processor.py` enthält bereits erste regelbasierte Logik für:
  - Duplikatprüfung
  - Kategorienpflicht
  - Schwierigkeitsgrad-Normalisierung
  - Pflichtfelder
- `src/main.py` nutzt Regelwerte bereits teilweise in Dialogen und Exportblockaden.

## Zielbild Architektur (Sprint 2)

## 1) Regelmodell erweitern

`data/config/import_rules.json` soll fachlich lesbar und modular aufgebaut sein.

### Zielbereiche

- `category_rules`
- `difficulty_rules`
- `template_rules`
- `duplicate_rules` bzw. bestehende Duplikat-Parameter
- `preview_rules`
- `field_alias_rules`
- `collection_governance`

## 2) Einheitliche Regelabfrage

Alle relevanten Entscheidungen sollen über zentrale Getter laufen:

- `WordProcessor.get_import_rule(...)`
- GUI-seitige Nutzung ohne parallele Sonderlogik
- Keine hart verdrahteten Fachwerte in Dialogtexten oder Exportprüfungen, wenn konfigurierbar

## 3) Fachliche Erweiterungen

### Konfigurierbar machen

- Pflichtfelder je Sammlungsmodus
- Alias-Mapping für strukturierte Tabellenfelder
- Behandlung fehlender optionaler Blöcke in der Vorschau
- Exportblockaden bei Pflichtfeld-/Kategorie-/Difficulty-Verstößen
- Standard-Metadaten für neue Aufgabenimporte

## Konkrete Umsetzungsschritte (Reihenfolge)

1. **Ist-Regeln inventarisieren**
   - Alle aktuell im Code vorhandenen Fachregeln erfassen
   - Harte Defaultwerte identifizieren

2. **Regeldatei strukturieren**
   - `import_rules.json` klarer gliedern
   - Fehlende Bereiche ergänzen

3. **Code auf zentrale Regelzugriffe umstellen**
   - `word_processor.py`
   - `main.py`
   - ggf. `template_manager.py`

4. **Alias- und Pflichtfeldregeln konsolidieren**
   - Tabellenmodus
   - H1/H2-Modus

5. **Doku nachziehen**
   - Technikdoku
   - Governance-Kurzregeln in `memos/`

## Akzeptanzkriterien (Sprint 2)

- Fachliche Kernregeln sind zentral konfigurierbar.
- GUI und Exportpfad verwenden dieselben Regelquellen.
- Neue Feldalias-Varianten lassen sich ohne Codeänderung ergänzen.
- Pflichtfeld- und Blockierlogik ist transparent dokumentiert.
- Keine neuen Probleme in der Probleme-Konsole.

## Testfälle (Minimum)

1. Kategoriepflicht per Regel aktiv/deaktiviert → Verhalten ändert sich konsistent.
2. Neues Alias für Intro/Hinweis in JSON ergänzt → Parser/Export erkennen es ohne Codeänderung.
3. Pflichtfeldsatz geändert → Diagnostik und Exportblockade folgen der Konfiguration.
4. Default-Metadaten im Import geändert → GUI übernimmt neue Defaultwerte.

## Risiken + Gegenmaßnahmen

- Risiko: Regeln werden zu komplex und schwer verständlich.
  - Gegenmaßnahme: klare Abschnittsstruktur + kommentierte Defaults.
- Risiko: GUI und Backend driften trotz Konfiguration auseinander.
  - Gegenmaßnahme: nur zentrale Getter nutzen, keine duplizierten Fachkonstanten.
- Risiko: Rückwärtskompatibilität bestehender Sammlungen leidet.
  - Gegenmaßnahme: stabile Defaults und additive Erweiterung.

## Deliverables Sprint 2

- Erweiterte `data/config/import_rules.json`
- Konsolidierte Regelabfragen in `src/word_processor.py` und `src/main.py`
- Aktualisierte technische Doku
- Kurze Governance-Notizen für neue Sammlungen
