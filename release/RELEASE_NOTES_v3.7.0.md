# Release Notes v3.7.0
<!-- markdownlint-disable MD012 -->

Datum: 2026-06-06

## Highlights

- Wizard-/Listen-UX ausgebaut:
  - **Nur Blocker anzeigen** + **Nur Blocker auswählen**
  - To-fix-Hinweise im Wizard-Status
  - Sortierbare Aufgabenliste (Spaltenheader)
  - Doppelklick-Bearbeitung einzelner Aufgaben
  - Mehrfachbearbeitung (Kategorie, Schwierigkeit, Schlagworte ersetzen/anhängen)
- Modussteuerung in der GUI:
  - Modus **einfach (empfohlen)** vs. **erweitert**
  - erweiterte Funktionen werden im einfachen Modus ausgeblendet
- Aufgabenübernahme/Export verbessert:
  - Übernahme von Tabellen und Inhaltssteuerelementen (z. B. Kontrollkästchen)
  - Aufgabentitel als `Überschrift 1`
  - Punkte am Titelende rechtsbündig mit Rahmen (konfigurierbar)
- Konfiguration und Hilfe erweitert:
  - `data/config/import_rules.json` um `export_rules.title_points_box.*`
  - GUI-Hilfe verweist explizit auf Export-Layout-Keys
- Versions- und Vorlagenpflege:
  - GUI-Titelleiste zeigt Version aus `build_version_info.txt`
  - Fallback ergänzt: Wenn `build_version_info.txt` im Deploy fehlt, wird die Version aus EXE-/Ordnernamen (z. B. `LEK-Bastler_3.7.0.exe`) ermittelt
  - Vorlagen bereinigt: `AUFGABEN-Vorlage.docx` + `LEK-Vorlage.docx`

## Qualitätsstatus

- Problems-Check ohne offene Diagnosen in den geänderten Dateien.
- Regressionstest-Suite: 19/19 grün (`tools/test_regression_core.py`).
- Release-Build für `3.7.0` aktualisiert (Artefakte siehe unten).

## Artefakte

- EXE: `dist/LEK-Bastler_3.7.0/LEK-Bastler_3.7.0.exe`
- Release-ZIP: `release/LEK-Bastler_3.7.0.zip`
- Release Notes: `release/RELEASE_NOTES_v3.7.0.md`
<!-- EOF -->
<!-- END -->


