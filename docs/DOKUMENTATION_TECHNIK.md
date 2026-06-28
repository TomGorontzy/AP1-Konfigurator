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

- `src/build.ps1` – erstellt die Portable-EXE per PyInstaller und paketiert das Release-ZIP
- `src/setup.ps1` – richtet die Python-Buildumgebung für die EXE-Erzeugung ein
- `src/publish_release.ps1` – veröffentlicht die EXE-Variante auf GitHub
- `src/AP1-Konfigurator.ps1` – Haupteinstieg und Orchestrierung
- `src/AP1-Konfigurator.bat` – interaktiver Starter mit Proxy-Abfrage
- `src/main.py` – EXE-Launcher, der eingebettete Laufzeitdateien nach `%LOCALAPPDATA%` synchronisiert, einen `current`-Arbeitsordner pflegt und von dort startet
- `src/post_build.py` – erstellt den versionierten Portable-Ordner und das ZIP-Artefakt
- `.github/workflows/release.yml` – GitHub-Workflow für tagbasierte EXE-Releases
- `src/Skript-Module/AP1-Logging.psm1` – Logging- und Konsolenausgabe
- `src/Skript-Module/AP1-Office.psm1` – First-Run-Erkennung und Office-Initialisierung
- `src/Skript-Module/AP1-Registry.psm1` – Registry-Helfer, COM-Bereinigung, Prüfungen
- `src/Skript-Module/AP1-Folders.psm1` – Kandidatenordner und Desktop-Bereitstellung
- `src/Skript-Module/AP1-Templates.psm1` – Vorlagen und Schnellzugriff
- `src/Skript-Module/AP1-System.psm1` – Proxy, Taskleiste, Prozesse, Archiv-Entpackung
- `src/Skript-Module/AP1-Nuera.psm1` – Download und Bereitstellung der Nuera-Dateien
- `src/Skript-Module/AP1-QuickAccess.psm1` – Schnellzugriff und Desktop-Shortcuts
- `data/1. Anpassen/` – anwendungsnahe Eingabedaten, insbesondere `AP1-TN.xlsx`
- `data/2. Bei Bedarf anpassen/` – Vorlagen, Office-UI-Dateien, temporäre Ordnerstruktur
- `data/3. Nuera-Dateien/` – Download- und Entpackbereich für Nuera-Dateien
- `data/4. Logs/` – Transcript-Ausgaben
- `docs/` – kanonische Projektdokumentation
- `release/`, `dist/`, `build/` – historische bzw. erzeugte Release-Artefakte

## Einstiegspunkt und Parameter

Einstiegspunkt:

- `src/AP1-Konfigurator.ps1`

Unterstützte Parameter:

| Parameter | Typ | Standard | Beschreibung |
| --- | --- | --- | --- |
| `-Proxy` | `On`, `Off`, `Skip` | `Skip` | Aktiviert, deaktiviert oder überspringt Proxy-Änderungen |
| `-ProxyServer` | `String` | `192.168.0.1:8080` | Proxy-Server für `-Proxy On` |
| `-ProxyBypass` | `String` | Office-/Microsoft-Ausnahmeliste | Proxy-Ausnahmen |
| `-ExcelListPath` | `String` | automatisch `data/1. Anpassen\AP1-TN.xlsx` | Pfad zur Teilnehmerliste |
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

1. Die EXE synchronisiert eingebettete Laufzeitdateien nach `%LOCALAPPDATA%\AP1-Konfigurator-Portable\vX.Y.Z`
2. Die EXE synchronisiert `data/` und `docs/` aus dem Release-Verzeichnis in denselben lokalen Versionsordner
3. Die EXE aktualisiert zusätzlich `%LOCALAPPDATA%\AP1-Konfigurator-Portable\current` als aktuelle Arbeitskopie
4. Die EXE entfernt ältere `%LOCALAPPDATA%\AP1-Konfigurator-Portable\v*`-Ordner automatisch
5. Modulimport aus `src/Skript-Module` (im Release-Laufzeitverzeichnis dann `Skript-Module`)
6. Setzen des Word-Templatepfads auf den Desktop des aktuellen Benutzers
7. Anheften des Desktops an den Schnellzugriff
8. Start des Transcript-Loggings in `data/4. Logs`
9. Prüfung der COM-Verfügbarkeit von Word und Excel
10. Bereitstellung der aktuellen Nuera-Dateien
11. Office-/Explorer-Konfiguration via Registry und optional COM
12. Kopieren von Vorlagen und Schnellzugriff-Dateien
13. Erzeugung der Kandidatenordner aus Excel oder CSV
14. Anwenden von Taskleisten- und Proxy-Einstellungen

## Besondere Laufzeitlogik

### COM-Fallback

- Standardmäßig wird Word/Excel per COM geprüft und verwendet.
- Wenn die COM-Prüfung fehlschlägt, wird automatisch in den Registry-Fallback gewechselt.
- Mit `-RegistryOnly` kann dieser Modus direkt erzwungen werden.

### Excel-/CSV-Fallback

- Primärquelle ist `data/1. Anpassen\AP1-TN.xlsx`.
- Wenn Excel/COM nicht verfügbar ist, kann mit `-CsvFallbackPath` auf CSV ausgewichen werden.
- Erwartet werden verwertbare Werte in den ersten beiden Spalten bzw. CSV-Feldern `Account` und `Kandidat`.

### Nuera-Dateien

- Es wird eine priorisierte Liste geprüft, aktuell beginnend mit `nuera2026_f.zip`.
- Nach erfolgreichem Download wird das Archiv entpackt und der extrahierte Ordner auf den Desktop kopiert.
- Der Download erfolgt direkt aus dem AkA-Bereich.

### Desktop-Bereitstellung

- Kandidaten- und Nuera-Ordner werden gezielt auf dem Desktop des aktuellen Benutzers abgelegt.
- Temporäre Ordner unter `data/2. Bei Bedarf anpassen\Ordner` werden im Anschluss aufgeräumt.

### EXE-Laufzeitverzeichnis

- Die PowerShell-Startdateien und `Skript-Module` werden nicht separat im Release abgelegt.
- Sie sind in der EXE eingebettet und werden beim Start nach `%LOCALAPPDATA%\AP1-Konfigurator-Portable\vX.Y.Z` kopiert.
- `data/` und `docs/` aus dem Release-Verzeichnis werden ebenfalls in dieses lokale Arbeitsverzeichnis synchronisiert.
- Zusätzlich wird `%LOCALAPPDATA%\AP1-Konfigurator-Portable\current` als aktuelle Arbeitskopie gepflegt und bevorzugt gestartet.
- Ältere versionierte Unterordner `v*` werden beim Start automatisch entfernt, sofern sie nicht der aktuell laufenden Version entsprechen.
- Das veröffentlichte EXE-Release enthält daher nur `AP1-Konfigurator-Portable.exe`, `data/`, `docs/` und `README.md`.
- Die erzeugte ZIP-Datei ist flach und enthält keinen Sammelordner als oberste Ebene.

## Logging

- Transcript-Dateien liegen unter `data/4. Logs`.
- Namensschema: `AP1_Prep_yyyyMMdd_HHmmss.log`
- Es bleiben maximal fünf Logs erhalten.
- `Write-SafeOutput` wird für sichtbare Statusmeldungen verwendet.

## Betriebs- und Wartungshinweise

- Vorlagen und Office-UI-Dateien liegen unter `data/2. Bei Bedarf anpassen/`.
- Office-abhängige Abläufe sollten idealerweise auf einem System mit funktionierender Word-/Excel-Installation getestet werden.
- Bei First-Run-Dialogen kann ein manueller Office-Start erforderlich sein.
- Historische Build- und Release-Artefakte im Repository sind nicht die kanonische Quellstruktur.
- Die kanonische Projektdokumentation liegt unter `docs/`.

## Release- und GitHub-Hinweise

- Aktueller veröffentlichter Stand im Repository: `v1.0.12`
- Changelog: `docs/CHANGELOG.md`
- Release-Notizen: `release/RELEASE_NOTES_v1.0.12.md`
- Der GitHub-Release wird über die EXE-Variante `AP1-Konfigurator-Portable-vX.Y.Z.zip` veröffentlicht.
- Das EXE-Release enthält nur `AP1-Konfigurator-Portable.exe`, `data/`, `docs/` und `README.md`.
- Für konsistente Releases sollten Anwender-, Technik-, Kurz- und Release-Prozess-Dokumentation vor dem Tagging aktualisiert werden.
