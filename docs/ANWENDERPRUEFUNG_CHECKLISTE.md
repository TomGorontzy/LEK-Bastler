# Anwenderprüfungs-Checkliste

## Ziel

Diese Checkliste unterstützt nicht-technische Prüferinnen und Prüfer dabei, die bereitgestellte Windows-EXE des LEK-Bastler strukturiert auf Funktionalität, Verständlichkeit und Bedienbarkeit zu testen.

## Für wen diese Checkliste gedacht ist

- Anwenderinnen und Anwender ohne Entwicklerwissen
- Kolleginnen und Kollegen aus dem Fachbereich
- Testpersonen, die mit echten Aufgabensammlungen arbeiten

## Vorbereitung

Vor dem Test bitte sicherstellen:

- Das Windows-Release wurde vollständig entpackt.
- Die EXE lässt sich lokal von einem normalen Arbeitsordner aus starten.
- Mindestens eine echte oder realistische Beispieldatei unter `data/Aufgaben/` ist vorhanden.
- Falls verfügbar, liegt auch eine schwierigere Testdatei mit Sonderfällen vor (z. B. Intro-Text, Tabellen, unvollständige Angaben).

## So soll getestet werden

Bitte nicht nur „klickt irgendwie“ bewerten, sondern bewusst auf folgende Fragen achten:

- Ist der nächste Schritt jeweils klar erkennbar?
- Sind Bezeichnungen und Hinweise verständlich?
- Wirkt das Verhalten zuverlässig und nachvollziehbar?
- Entspricht das Ergebnis dem, was vorher in der Anwendung sichtbar war?

## 1. Start und erster Eindruck

Bitte prüfen:

- Lässt sich die EXE ohne technische Zusatzschritte starten?
- Ist nach dem Start sofort erkennbar, was zu tun ist?
- Wirkt die Oberfläche übersichtlich und verständlich?
- Sind Buttons, Eingabefelder und Hinweise sprachlich klar?

Bewertung:

- [ ] unauffällig / verständlich
- [ ] teilweise unklar
- [ ] deutlich verbesserungsbedürftig

Notizen:

- ...

## 2. Datei laden

Bitte prüfen:

- Ist das Laden einer Aufgabensammlung einfach auffindbar?
- Ist klar, welche Datei ausgewählt werden soll?
- Reagiert die Anwendung nachvollziehbar nach dem Laden?
- Werden bei problematischen Dateien verständliche Hinweise angezeigt?

Bewertung:

- [ ] unauffällig / verständlich
- [ ] teilweise unklar
- [ ] deutlich verbesserungsbedürftig

Notizen:

- ...

## 3. Aufgabenerkennung und Sichtprüfung

Bitte prüfen:

- Werden die erkannten Aufgaben plausibel angezeigt?
- Werden einleitende Texte oder Hinweise nicht fälschlich als Aufgabe behandelt?
- Sind wichtige Informationen zur Aufgabe verständlich dargestellt?
- Fällt auf, wenn etwas fehlt, doppelt erscheint oder offensichtlich falsch zugeordnet wurde?

Bewertung:

- [ ] unauffällig / verständlich
- [ ] teilweise unklar
- [ ] deutlich verbesserungsbedürftig

Notizen:

- ...

## 4. Auswahl, Freigabe und Navigation

Bitte prüfen:

- Ist klar, wie Aufgaben ausgewählt oder freigegeben werden?
- Ist verständlich, wie man zum nächsten Schritt gelangt?
- Kann man frühere Schritte nachvollziehbar erneut prüfen?
- Wirkt die Reihenfolge der Arbeitsschritte logisch?

Bewertung:

- [ ] unauffällig / verständlich
- [ ] teilweise unklar
- [ ] deutlich verbesserungsbedürftig

Notizen:

- ...

## 5. Vorschau

Bitte prüfen:

- Ist die Vorschau leicht zu verstehen?
- Entsteht Vertrauen, dass das spätere Ergebnis genauso aussieht bzw. inhaltlich dazu passt?
- Sind fehlende oder auffällige Inhalte gut erkennbar?
- Ist klar, ob noch eine Korrektur nötig ist oder exportiert werden kann?

Bewertung:

- [ ] unauffällig / verständlich
- [ ] teilweise unklar
- [ ] deutlich verbesserungsbedürftig

Notizen:

- ...

## 6. Export und Ergebnisdatei

Bitte prüfen:

- Ist der Export leicht auffindbar und verständlich auslösbar?
- Wird klar, wo die erzeugte Datei gespeichert wird?
- Entspricht die exportierte LEK inhaltlich der bestätigten Vorschau?
- Wirkt die erzeugte Datei vollständig und fachlich plausibel?

Bewertung:

- [ ] unauffällig / verständlich
- [ ] teilweise unklar
- [ ] deutlich verbesserungsbedürftig

Notizen:

- ...

## 7. Warnungen, Hinweise und Fehlermeldungen

Bitte prüfen:

- Sind Warnungen verständlich formuliert?
- Ist klar, was das Problem ist?
- Ist klar, was als Nächstes zu tun wäre?
- Wirken Sperren oder Blockaden nachvollziehbar?

Bewertung:

- [ ] unauffällig / verständlich
- [ ] teilweise unklar
- [ ] deutlich verbesserungsbedürftig

Notizen:

- ...

## 8. Gesamturteil zur Bedienbarkeit

Bitte einschätzen:

- Würden Sie die Anwendung ohne zusätzliche Erklärung erneut verwenden?
- Würden Sie einer Kollegin oder einem Kollegen die Nutzung zutrauen?
- An welchen Stellen war Unsicherheit, Nachdenken oder Raten nötig?

Gesamtbewertung:

- [ ] sofort alltagstauglich
- [ ] grundsätzlich brauchbar, aber mit kleineren Hürden
- [ ] nur mit zusätzlicher Erklärung gut nutzbar
- [ ] aktuell zu unklar oder fehleranfällig

## 9. Wichtigste Rückmeldungen

Bitte möglichst konkret notieren:

### Was hat gut funktioniert?

- ...

### Was war unklar oder irritierend?

- ...

### Wo sind Fehler oder fachliche Auffälligkeiten aufgefallen?

- ...

### Welche Verbesserung hätte den größten Nutzen?

- ...

## 10. Minimale Testabdeckung

Ein vollständiger Anwendercheck sollte mindestens Folgendes enthalten:

- Start der Windows-EXE
- Laden von mindestens einer realistischen Aufgabensammlung
- Sichtprüfung der erkannten Aufgaben
- Auswahl bzw. Freigabe einer Teilmenge
- Prüfung der Vorschau
- Export einer LEK
- Kurzbewertung von Verständlichkeit und Bedienbarkeit

## Ergebnis der Anwenderprüfung

- Datum:
- Name der Testperson:
- Verwendete Beispieldatei:
- Ergebnis:
  - [ ] ohne Auffälligkeiten
  - [ ] kleinere Probleme festgestellt
  - [ ] deutliche Probleme festgestellt
  - [ ] weitere Prüfung vor Einsatz empfohlen
