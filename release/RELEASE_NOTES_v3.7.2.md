# Release Notes v3.7.2
## Enthalten

- Stabilisierung der GUI-Versionsanzeige: Fallback auf EXE-/Ordnernamen, wenn `build_version_info.txt` im Deploy-Kontext fehlt.
- Erweiterte Aufgaben-Gruppenlogik im Wizard/Export: Haupt- und Unteraufgaben werden konsistent gemeinsam behandelt.
- Unterstützung externer Tabellen-Referenzen per Marker `<<tabelle=...>>` inkl. automatischer Hoch-/Querformat-Behandlung.
- Fix für Absatzübernahme aus Aufgabeninhalten:
  - Leere Zeilen bleiben nach der Aufgabenbearbeitung erhalten.
  - Fallback-Export übernimmt leere Zeilen als leere Word-Absätze.

## Qualitätssicherung

- Syntax-/Problems-Check der geänderten Python-Dateien ohne Fehler.
- Reproduktionstest für Leerabsatz-Fallback erfolgreich (Word-Ausgabe enthält Leerabsätze).

## Artefakte

- dist/LEK-Bastler_3.7.2/LEK-Bastler_3.7.2.exe
- release/LEK-Bastler_3.7.2.zip
- release/RELEASE_NOTES_v3.7.2.md

