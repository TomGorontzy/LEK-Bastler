# MEMO: Governance für neue Aufgabensammlungen

## Ziel

Neue Sammlungen sollen ohne Codeänderungen in den Import-/Exportfluss passen.

## Mindestregeln

- Neue Feldvarianten **nur** in `data/config/import_rules.json` unter `field_alias_rules.structured_task_fields` ergänzen.
- Feldbezeichnungen für Hinweise/Warnungen in `field_alias_rules.labels` pflegen.
- Pflichtfeldänderungen in `template_rules.required_fields` vornehmen.
- Export-Blockierverhalten ausschließlich über
  - `template_rules.block_export_on_missing_required`
  - `category_rules.block_export_on_missing`
  - `difficulty_rules.block_export_on_inconsistent`
  steuern.

## Reihenfolge für Vorschau/Export

- Reihenfolge und Optionalität der Blöcke in `preview_rules.task_flow_sections` definieren.
- Delta-Hinweise auf fehlende optionale Blöcke über
  - `preview_rules.show_optional_missing_sections`
  - `preview_rules.optional_sections`
  steuern.

## Änderungen an Sammlungsformaten

1. Alias in `field_alias_rules` ergänzen.
2. Beispiel-Datei importieren und Vorschau prüfen.
3. Export mit Pflichtprüfungen validieren.
4. Doku (`docs/DOKUMENTATION_TECHNIK.md`) bei strukturellen Änderungen aktualisieren.

## Nicht erlaubt

- Fachwerte „schnell“ hart im Code nachziehen, wenn derselbe Effekt über `import_rules.json` möglich ist.
- Neue Pflichtfeldprüfungen ohne Dokumentation und ohne Regelpfad einbauen.
