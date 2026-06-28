# DOKUMENTATION TECHNIK

## Inhaltsverzeichnis

- [Projektüberblick](#projektüberblick)
- [Repository-Struktur](#repository-struktur)
- [Einstiegspunkt und Parameter](#einstiegspunkt-und-parameter)
- [Module und Zuständigkeiten](#module-und-zuständigkeiten)
- [Ablauf der Ausführung](#ablauf-der-ausführung)
- [Besondere Laufzeitlogik](#besondere-laufzeitlogik)
- [Logging](#logging)
- [Betriebs- und Wartungshinweise](#betriebs--und-wartungshinweise)
- [Release- und GitHub-Hinweise](#release--und-github-hinweise)

## Projektüberblick

Der `AP1-Konfigurator` ist ein PowerShell-basiertes Automatisierungsskript zur Einrichtung von Prüfungsrechnern für die AP1.

Zentrale Aufgaben:

- Office-Initialisierung und Registry-Konfiguration
- Übernahme von Word-/Excel-Vorlagen
- Anlage von Kandidatenordnern aus Excel oder CSV
- Download und Bereitstellung der Nuera-Dateien
- optionale Proxy-Konfiguration
- Transcript-Logging für Nachvollziehbarkeit

## Repository-Struktur

- `AP1-Konfigurator.ps1` – Haupteinstieg und Orchestrierung
- `AP1-Konfigurator.bat` – interaktiver Starter mit Proxy-Abfrage
- `Skript-Module/AP1-Logging.psm1` – Logging- und Konsolenausgabe
- `Skript-Module/AP1-Office.psm1` – First-Run-Erkennung und Office-Initialisierung
- `Skript-Module/AP1-Registry.psm1` – Registry-Helfer, COM-Bereinigung, Prüfungen
- `Skript-Module/AP1-Folders.psm1` – Kandidatenordner und Desktop-Bereitstellung
- `Skript-Module/AP1-Templates.psm1` – Vorlagen und Schnellzugriff
- `Skript-Module/AP1-System.psm1` – Proxy, Taskleiste, Prozesse, Archiv-Entpackung
- `Skript-Module/AP1-Nuera.psm1` – Download und Bereitstellung der Nuera-Dateien
- `Skript-Module/AP1-QuickAccess.psm1` – Schnellzugriff und Desktop-Shortcuts
- `1. Anpassen/` – anwendungsnahe Eingabedaten, insbesondere `AP1-TN.xlsx`
- `2. Bei Bedarf anpassen/` – Vorlagen, Office-UI-Dateien, temporäre Ordnerstruktur
- `3. Nuera-Dateien/` – Download- und Entpackbereich für Nuera-Dateien
- `4. Logs/` – Transcript-Ausgaben
- `docs/` – kanonische Projektdokumentation
- `release/`, `dist/`, `build/` – historische bzw. erzeugte Release-Artefakte

## Einstiegspunkt und Parameter

Einstiegspunkt:

- `AP1-Konfigurator.ps1`

Unterstützte Parameter:

| Parameter | Typ | Standard | Beschreibung |
| --- | --- | --- | --- |
| `-Proxy` | `On`, `Off`, `Skip` | `Skip` | Aktiviert, deaktiviert oder überspringt Proxy-Änderungen |
| `-ProxyServer` | `String` | `192.168.0.1:8080` | Proxy-Server für `-Proxy On` |
| `-ProxyBypass` | `String` | Office-/Microsoft-Ausnahmeliste | Proxy-Ausnahmen |
| `-ExcelListPath` | `String` | automatisch `1. Anpassen\AP1-TN.xlsx` | Pfad zur Teilnehmerliste |
| `-CsvFallbackPath` | `String` | leer | CSV-Fallback, wenn COM/Excel nicht verwendbar ist |
| `-MaxRows` | `Int` | `500` | Obergrenze der zu lesenden Datensätze |
| `-Quiet` | `Switch` | aus | Unterdrückt die interaktive Proxy-Bestätigung |
| `-RegistryOnly` | `Switch` | aus | Erzwingt den Betrieb ohne COM |

## Module und Zuständigkeiten

| Modul | Verantwortung |
| --- | --- |
| `AP1-Logging.psm1` | Transcript-Logging, sichtbare Ausgaben, Log-Rotation |
| `AP1-Office.psm1` | First-Run-Erkennung, robuster Word-/Excel-Start |
| `AP1-Registry.psm1` | Registry-Schreibzugriffe, COM-Hilfsfunktionen, Tests |
| `AP1-Folders.psm1` | Erzeugung von Kandidatenordnern aus Excel/CSV |
| `AP1-Templates.psm1` | Kopieren von `Normal.dotm`, `Mappe.xltx`, Office-UI-Dateien |
| `AP1-System.psm1` | Proxy, Taskleiste, Prozessstopp, ZIP-Entpackung |
| `AP1-Nuera.psm1` | priorisierte Nuera-Downloadlogik und Desktop-Kopie |
| `AP1-QuickAccess.psm1` | Desktop an Schnellzugriff anheften |
| `AP1-Utils.psm1` | Hilfsfunktionen für Text und Verknüpfungen |

## Ablauf der Ausführung

1. Modulimport aus `Skript-Module`
2. Setzen des Word-Templatepfads auf den Desktop des aktuellen Benutzers
3. Anheften des Desktops an den Schnellzugriff
4. Start des Transcript-Loggings in `4. Logs`
5. Prüfung der COM-Verfügbarkeit von Word und Excel
6. Bereitstellung der aktuellen Nuera-Dateien
7. Office-/Explorer-Konfiguration via Registry und optional COM
8. Kopieren von Vorlagen und Schnellzugriff-Dateien
9. Erzeugung der Kandidatenordner aus Excel oder CSV
10. Anwenden von Taskleisten- und Proxy-Einstellungen

## Besondere Laufzeitlogik

### COM-Fallback

- Standardmäßig wird Word/Excel per COM geprüft und verwendet.
- Wenn die COM-Prüfung fehlschlägt, wird automatisch in den Registry-Fallback gewechselt.
- Mit `-RegistryOnly` kann dieser Modus direkt erzwungen werden.

### Excel-/CSV-Fallback

- Primärquelle ist `1. Anpassen\AP1-TN.xlsx`.
- Wenn Excel/COM nicht verfügbar ist, kann mit `-CsvFallbackPath` auf CSV ausgewichen werden.
- Erwartet werden verwertbare Werte in den ersten beiden Spalten bzw. CSV-Feldern `Account` und `Kandidat`.

### Nuera-Dateien

- Es wird eine priorisierte Liste geprüft, aktuell beginnend mit `nuera2026_f.zip`.
- Nach erfolgreichem Download wird das Archiv entpackt und der extrahierte Ordner auf den Desktop kopiert.
- Der Download erfolgt direkt aus dem AkA-Bereich.

### Desktop-Bereitstellung

- Kandidaten- und Nuera-Ordner werden gezielt auf dem Desktop des aktuellen Benutzers abgelegt.
- Temporäre Ordner unter `2. Bei Bedarf anpassen\Ordner` werden im Anschluss aufgeräumt.

## Logging

- Transcript-Dateien liegen unter `4. Logs`.
- Namensschema: `AP1_Prep_yyyyMMdd_HHmmss.log`
- Es bleiben maximal fünf Logs erhalten.
- `Write-SafeOutput` wird für sichtbare Statusmeldungen verwendet.

## Betriebs- und Wartungshinweise

- Vorlagen und Office-UI-Dateien liegen unter `2. Bei Bedarf anpassen/`.
- Office-abhängige Abläufe sollten idealerweise auf einem System mit funktionierender Word-/Excel-Installation getestet werden.
- Bei First-Run-Dialogen kann ein manueller Office-Start erforderlich sein.
- Historische Build- und Release-Artefakte im Repository sind nicht die kanonische Quellstruktur.
- Die kanonische Projektdokumentation liegt unter `docs/`.

## Release- und GitHub-Hinweise

- Aktueller veröffentlichter Stand im Repository: `v1.0.11`
- Changelog: `CHANGELOG.md`
- Release-Notizen: `RELEASE_NOTES_v1.0.11.md`
- Der GitHub-Release wird über Tag + `gh release create` veröffentlicht.
- Für konsistente Releases sollten Anwender-, Technik- und Kurzdokumentation vor dem Tagging aktualisiert werden.
