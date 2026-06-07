# Release Notes v3.7.5

## Enthalten

- Externe Tabellenmarker im strukturierten Export robust korrigiert:
  - Kombinationen wie `<<tabelle=Entgeltabrechnung.docx>>` und `<<tabelle_format=portrait>>` erzeugen keinen Doppelinsert mehr.
  - Der Exportpfad erkennt, ob externe Inhalte bereits eingefügt wurden, und unterbindet den Fallback-Doppeleintrag.
- Übernahme von Schrift-/Stilkontext aus externen Dokumenten verbessert:
  - Style-ID-Konflikte (z. B. `Normal`) werden konfliktfrei in `ext_*`-Styles überführt.
  - Implizite Default-Absatzstile werden explizit gesetzt und korrekt auf remappte Styles abgebildet.
- GUI-Resize-Stabilisierung:
  - Der Bereich unterhalb von „Import-Assistent“ bleibt bei Fensteränderungen stabil (kein Flattern durch Layout-Rückkopplung).

## Qualitätssicherung

- Regressionstest-Suite erfolgreich: **32/32 OK** (`tools/test_regression_core.py`).

## Artefakte

- dist/LEK-Bastler_3.7.5/LEK-Bastler_3.7.5.exe
- release/LEK-Bastler_3.7.5.zip
- release/RELEASE_NOTES_v3.7.5.md
