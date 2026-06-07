# Release Notes v3.7.4

## Enthalten

- Buildprozess um eine verpflichtende Markdown-Lintprüfung erweitert:
  - `src/build.ps1` führt standardmäßig `markdownlint-cli2` aus.
  - Bei Markdown-Verstößen wird der Build jetzt mit klarer Meldung gestoppt.
  - Optional kann der Schritt mit `-SkipLint` gezielt übersprungen werden.
- GUI-Aktionsbereich kompakter und übersichtlicher angeordnet:
  - `Auswahl` und `Prüfung` links untereinander,
  - `Export` rechtsbündig,
  - adaptive Breitenanpassung bei Fenster-Resize.
- Validierung für doppelte Aufgabennummern verschärft:
  - Erkennung beim Laden,
  - blockierende Warnungen,
  - klare Aufforderung zur Korrektur vor Export.
- Markdown-Dokumentation bereinigt, damit Lint im Build stabil ohne Fehler durchläuft.

## Qualitätssicherung

- Regressionstest-Suite erfolgreich: **30/30 OK** (`tools/test_regression_core.py`).
- Build mit aktivem Lint-Schritt erfolgreich (`src/build.ps1 -SkipBuild`), `markdownlint-cli2`: **0 Fehler**.

## Artefakte

- dist/LEK-Bastler_3.7.4/LEK-Bastler_3.7.4.exe
- release/LEK-Bastler_3.7.4.zip
- release/RELEASE_NOTES_v3.7.4.md
