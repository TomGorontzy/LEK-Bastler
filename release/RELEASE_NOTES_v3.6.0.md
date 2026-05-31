# Release Notes v3.6.0

Datum: 2026-05-27

## Highlights

- Sprint 2: Zentrales Regelwerk in `data/config/import_rules.json` ausgebaut
  - Alias-Regeln (`field_alias_rules`)
  - Vorschau-/Export-Regeln (`preview_rules`)
  - Governance-Hinweise (`collection_governance`)
- Sprint 3: Parser-Refactor in `src/word_processor.py`
  - klare Trennung von Strukturerkennung, Moduswahl/Fallback und Normalisierung
  - konfigurierbare Parser-Modi via `parser_rules`
- Sprint 4: Regressionstest- und QA-Rahmen
  - neue Test-Suite `tools/test_regression_core.py`
  - Test-Matrix `memos/MEMO_REGRESSIONSTEST_MATRIX.md`
  - Release-Checkliste `docs/RELEASE_QA_CHECKLISTE.md`

## Qualitätsstatus

- Regressionstest-Suite: 6/6 grün
- Parser-Mindesttests: 5/5 grün
- Smoke-Test: Blockade- und Erfolgsfall verifiziert

## Artefakte

- EXE: `dist/LEK-Bastler_3.6.0/LEK-Bastler_3.6.0.exe`
- Geplantes Release-Paket: `release/LEK-Bastler_3.6.0.zip`
- Release Notes: `release/RELEASE_NOTES_v3.6.0.md`
