# DOKUMENTATION ANWENDER

## Inhaltsverzeichnis

- [Zweck](#zweck)
- [Voraussetzungen](#voraussetzungen)
- [Vorbereitung](#vorbereitung)
- [Start des Skripts](#start-des-skripts)
- [Ablauf während der Ausführung](#ablauf-während-der-ausführung)
- [Ergebnis prüfen](#ergebnis-prüfen)
- [Häufige Probleme](#häufige-probleme)
- [Support-Hinweis](#support-hinweis)

## Zweck

Der `AP1-Konfigurator` richtet einen Prüfungsrechner für die Abschlussprüfung Teil 1 (AP1) standardisiert ein.

Dabei werden insbesondere vorbereitet:

- Microsoft Word und Excel
- Office-Vorlagen und Schnellzugriff
- Kandidatenordner aus der Teilnehmerliste
- Nuera-Dateien auf dem Desktop
- optionale Proxy-Einstellungen
- ein Laufzeitprotokoll zur Nachvollziehbarkeit

## Voraussetzungen

- Windows-PC mit Microsoft Office (Word und Excel)
- Schreibrechte im Benutzerprofil
- vorhandene Teilnehmerliste unter `data/1. Anpassen\AP1-TN.xlsx`
- optional Internetzugang für den Abruf der Nuera-Dateien

## Vorbereitung

1. Prüfen Sie, ob `AP1-TN.xlsx` im Ordner `data/1. Anpassen` vorhanden und aktuell ist.
2. Schließen Sie Word und Excel vor dem Start.
3. Speichern Sie offene Arbeiten anderer Programme.
4. Entscheiden Sie, ob ein Proxy gesetzt werden soll.
5. Halten Sie den Desktop des aktuellen Benutzers frei, damit die erzeugten Ordner sichtbar bleiben.

## Start des Skripts

Empfohlener Start:

1. `src/AP1-Konfigurator.bat` per Doppelklick öffnen.
2. Proxy-Frage beantworten.
3. Warten, bis das Skript die Einrichtung abgeschlossen hat.

Alternativ per PowerShell:

- `./src/AP1-Konfigurator.ps1`
- `./src/AP1-Konfigurator.ps1 -Proxy Off`
- `./src/AP1-Konfigurator.ps1 -RegistryOnly -CsvFallbackPath .\data\1. Anpassen\AP1-TN.csv`

## Ablauf während der Ausführung

Das Skript führt typischerweise folgende Schritte aus:

1. Office-Umgebung vorbereiten
2. Desktop an Schnellzugriff anheften
3. neueste Nuera-Dateien ermitteln und bereitstellen
4. Standardpfade für Word und Excel auf den Desktop setzen
5. Vorlagen `Normal.dotm` und `Mappe.xltx` kopieren
6. Kandidatenordner aus der Teilnehmerliste erzeugen
7. Taskleisten- und optional Proxy-Einstellungen anwenden
8. Logdatei in `data/4. Logs` schreiben

Hinweise während der Ausführung:

- Beim ersten Office-Start können Lizenz- oder Datenschutzhinweise erscheinen.
- Wenn Word oder Excel nicht per COM gestartet werden können, arbeitet das Skript teilweise im Registry-Fallback weiter.
- Die Ausführung kann je nach Office-Status und Netzwerkverbindung einige Minuten dauern.

## Ergebnis prüfen

Nach erfolgreicher Ausführung sollten Sie insbesondere Folgendes prüfen:

- gewünschter Nuera-Ordner liegt auf dem Desktop
- Kandidatenordner wurde auf dem Desktop angelegt
- Word/Excel-Speicherpfade zeigen auf den Desktop
- Schnellzugriff in Word/Excel wurde übernommen
- im Ordner `data/4. Logs` wurde eine aktuelle Logdatei angelegt

## Häufige Probleme

### Excel-Datei fehlt

- Prüfen, ob `data/1. Anpassen\AP1-TN.xlsx` vorhanden ist.
- Falls die Datei an anderem Ort liegt, `-ExcelListPath` verwenden.

### Word oder Excel startet nicht automatisch

- Office einmal manuell öffnen und Hinweise bestätigen.
- Danach das Skript erneut starten.
- Falls nötig mit `-RegistryOnly` arbeiten.

### Kandidatenordner wird nicht angelegt

- Teilnehmerliste auf Inhalte in Spalte A und B prüfen.
- Bei Excel-Problemen optional einen CSV-Fallback verwenden.

### Proxy-Rückfrage stört einen automatisierten Lauf

- Skript mit `-Quiet` starten.

### Nuera-Dateien werden nicht geladen

- Internetverbindung prüfen.
- Falls der Downloadserver nicht erreichbar ist, später erneut versuchen.

## Support-Hinweis

Bei Rückfragen bitte immer mitgeben:

- Datum/Uhrzeit der Ausführung
- Name der erzeugten Logdatei aus `data/4. Logs`
- kurze Beschreibung des letzten sichtbaren Schritts
- wenn möglich Screenshot einer Fehlermeldung
