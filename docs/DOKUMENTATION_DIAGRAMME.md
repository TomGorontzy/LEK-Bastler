# DOKUMENTATION_DIAGRAMME

Dieses Verzeichnis bündelt die ausgelagerten Mermaid-Quellen und die daraus erzeugten SVG-Grafiken für den LEK-Bastler.

## Speicherort

- `docs/diagramme/`

## Enthaltene Dateien

### Grafikdateien (SVG)

- `docs/diagramme/anwender_ablauf.svg`
- `docs/diagramme/technik_systemuebersicht.svg`
- `docs/diagramme/release_pipeline.svg`

### Quell-Dateien (Mermaid)

- `docs/diagramme/anwender_ablauf.mmd`
- `docs/diagramme/technik_systemuebersicht.mmd`
- `docs/diagramme/release_pipeline.mmd`

## Einbindung in die Dokumentation

- Anwenderfluss: `docs/DOKUMENTATION_ANWENDER.md`
- Techniküberblick: `docs/DOKUMENTATION_TECHNIK.md`
- Release-Ablauf: `docs/DOKUMENTATION_RELEASES.md`

## Empfohlene Positionierung (aktueller Stand)

- `anwender_ablauf.svg`
  - Position: direkt nach Abschnitt **Typischer Ablauf** in `docs/DOKUMENTATION_ANWENDER.md`
  - Zweck: Schrittfolge Auswahl → Freigabe → Export visuell verdichten
- `technik_systemuebersicht.svg`
  - Position: direkt nach Abschnitt **Pfadkonzept/Systemübersicht** in `docs/DOKUMENTATION_TECHNIK.md`
  - Zweck: Modul- und Datenflüsse inkl. Wizard/ImportSession auf einen Blick
- `release_pipeline.svg`
  - Position: direkt nach Abschnitt **GitHub Release-Workflow** in `docs/DOKUMENTATION_RELEASES.md`
  - Zweck: Build-, QA- und Publishing-Kette transparent machen

## Visuelle Legende (einheitlich)

Die Mermaid-Quellen nutzen ein gemeinsames Farbschema und Klassenmodell:

- `ui` = Benutzeroberfläche / Steuerung
- `process` = Verarbeitung / Modul-Logik
- `quality` = Prüfung / Freigabe / QA
- `decision` = Entscheidung / Gating
- `data` = Dateien, Ordner, Konfigurationen
- `release` = Build-/Release-Artefakte

## Starter-Snippet für neue Diagramme

Für neue Mermaid-Diagramme kann dieses kompakte Grundgerüst direkt übernommen werden:

```mermaid
flowchart LR
    A[Beispiel UI] --> B[Beispiel Prozess]
    B --> C{Entscheidung}
    C -->|Ja| D[Qualität/Prüfung]
    C -->|Nein| E[Datenziel]
    D --> F[Release-Artefakt]

    classDef ui fill:#DCEEFF,stroke:#1565C0,color:#0B3D91,stroke-width:1.4px;
    classDef process fill:#FFF2CC,stroke:#EF6C00,color:#8A3D00,stroke-width:1.4px;
    classDef data fill:#DFF3E3,stroke:#2E7D32,color:#1B5E20,stroke-width:1.4px;
    classDef quality fill:#F1E4FF,stroke:#7B1FA2,color:#4A148C,stroke-width:1.4px;
    classDef decision fill:#FFE3E3,stroke:#C62828,color:#8E0000,stroke-width:1.4px;
    classDef release fill:#E8E3FF,stroke:#512DA8,color:#311B92,stroke-width:1.4px;

    class A ui;
    class B process;
    class C decision;
    class D quality;
    class E data;
    class F release;

    linkStyle default stroke:#546E7A,stroke-width:1.9px;
```

Hinweis: `linkStyle` mit `stroke` und `stroke-width` ist bewusst mermaid-cli-kompatibel gehalten.

### Mini Do/Don’t (Render-stabil)

- ✅ Do: In `linkStyle` nur `stroke` und `stroke-width` verwenden.
- ✅ Do: Einrückungen mit Spaces statt Tabs schreiben (auch im Mermaid-Codeblock).
- ✅ Do: Klassen (`classDef`) zentral definieren und per `class ...` zuweisen.
- ❌ Don’t: `color` in `linkStyle` nutzen (führt in mermaid-cli häufig zu Parse-Fehlern).
- ❌ Don’t: Mischformatierung aus Tabs/Spaces verwenden.
- ❌ Don’t: Nicht validierte Stilattribute „auf Verdacht“ einführen.

## Hinweis

Die Mermaid-Dateien sind die pflegbaren Quellen. Die SVG-Dateien dienen als direkt einbindbare Grafiken für GitHub, VS Code und Release-Dokumentation.
