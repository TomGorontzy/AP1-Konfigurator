
# AP1-Konfigurator

Aktueller Stand: **v1.0.11** · Letzte Aktualisierung: **28. Juni 2026**

Der `AP1-Konfigurator` automatisiert die Einrichtung von Prüfungsrechnern für die Abschlussprüfung Teil 1 (AP1). Das PowerShell-Skript richtet Office, Explorer, Schnellzugriff, Proxy-Einstellungen, Kandidatenordner und die Nuera-Dateien in einer reproduzierbaren Reihenfolge ein.

## Was das Skript erledigt

- initialisiert Word und Excel robust über COM mit Registry-Fallback
- setzt Standard-Speicherpfade auf den Desktop des aktuellen Benutzers
- übernimmt `Normal.dotm`, `Mappe.xltx` und Office-Schnellzugriff
- erzeugt Kandidatenordner aus `AP1-TN.xlsx` oder optional aus einer CSV-Datei
- lädt die neuesten Nuera-Dateien und kopiert sie auf den Desktop
- setzt Taskleisten- und optionale Proxy-Einstellungen
- schreibt Laufzeitprotokolle nach `4. Logs`

## Schnellstart

Interaktiv per Batch:

```powershell
.\AP1-Konfigurator.bat
```

Direkt per PowerShell:

```powershell
.\AP1-Konfigurator.ps1
```

## Verfügbare Parameter

| Parameter | Typ | Standard | Zweck |
| --- | --- | --- | --- |
| `-Proxy` | `On`, `Off`, `Skip` | `Skip` | Steuert die Proxy-Konfiguration |
| `-ProxyServer` | `String` | `192.168.0.1:8080` | Proxy-Server bei `-Proxy On` |
| `-ProxyBypass` | `String` | Office-/Microsoft-Bypassliste | Ausnahmeliste für den Proxy |
| `-ExcelListPath` | `String` | automatisch `1. Anpassen\AP1-TN.xlsx` | Teilnehmerliste für die Ordnererzeugung |
| `-CsvFallbackPath` | `String` | leer | CSV-Fallback, wenn Excel/COM nicht verfügbar ist |
| `-MaxRows` | `Int` | `500` | Obergrenze für gelesene Teilnehmerzeilen |
| `-Quiet` | `Switch` | aus | Unterdrückt die Proxy-Rückfrage |
| `-RegistryOnly` | `Switch` | aus | Erzwingt den Betrieb ohne COM |

## Typische Aufrufe

```powershell
.\AP1-Konfigurator.ps1 -Proxy Off
.\AP1-Konfigurator.ps1 -Proxy On -Quiet
.\AP1-Konfigurator.ps1 -RegistryOnly -CsvFallbackPath .\1. Anpassen\AP1-TN.csv
```

## Ordnerstruktur

```text
AP1-Konfigurator/
├── AP1-Konfigurator.ps1
├── AP1-Konfigurator.bat
├── Skript-Module/
├── 1. Anpassen/
├── 2. Bei Bedarf anpassen/
├── 3. Nuera-Dateien/
├── 4. Logs/
├── docs/
├── dist/
├── release/
└── _Archive/
```

## Wichtige Hinweise

- Die Nuera-Dateien werden im aktuellen Stand **automatisch** geprüft und bereitgestellt.
- Wird Word oder Excel per COM nicht verfügbar, wechselt das Skript automatisch in den Registry-Fallback.
- Für einen stillen Proxy-Lauf ohne Rückfrage sollte `-Quiet` mitgegeben werden.
- Die Laufzeit hängt stark davon ab, ob Office erstmals gestartet werden muss.
- Releases sollen grundsätzlich über die **EXE-Variante** erfolgen (`AP1-Konfigurator-Portable`).
- Die EXE entpackt bzw. synchronisiert ihre eingebetteten Laufzeitdateien beim Start nach `%LOCALAPPDATA%\AP1-Konfigurator-Portable\vX.Y.Z` und verwendet diese Kopie für die Ausführung.

## EXE-Release-Struktur

Das Portable-Release ist bewusst schlank gehalten und enthält nur:

- `AP1-Konfigurator-Portable.exe`
- `data/`
- `docs/`
- `README.md`

Die PowerShell-Startskripte und `Skript-Module` werden nicht separat mit ausgeliefert, sondern befinden sich in der EXE und werden beim Start nach `%LOCALAPPDATA%` kopiert.

## Weiterführende Doku

- Anwender-Dokumentation: [`docs/DOKUMENTATION_ANWENDER.md`](./docs/DOKUMENTATION_ANWENDER.md)
- Technische Details: [`docs/DOKUMENTATION_TECHNIK.md`](./docs/DOKUMENTATION_TECHNIK.md)
- Kurzinfo: [`docs/KURZDOKUMENTATION.txt`](./docs/KURZDOKUMENTATION.txt)
- Checkliste: [`docs/DOKUMENTATION_CHECKLISTE.md`](./docs/DOKUMENTATION_CHECKLISTE.md)
- Release-Prozess: [`docs/RELEASE_PROZESS.md`](./docs/RELEASE_PROZESS.md)
- Änderungen: [`CHANGELOG.md`](./CHANGELOG.md)
- Release-Hinweise: [`RELEASE_NOTES_v1.0.11.md`](./RELEASE_NOTES_v1.0.11.md)
