
# AP1-Konfigurator – technische Dokumentation

Stand: **v1.0.10** · Aktualisiert am **28. Juni 2026**

## Zweck

Der `AP1-Konfigurator` richtet Prüfungsrechner für die AP1 standardisiert ein. Der Fokus liegt auf einer stabilen Office- und Desktop-Konfiguration, der Bereitstellung der Teilnehmerordner und der Ablage der aktuellen Nuera-Dateien.

## Ablauf der Ausführung

1. Laden aller Module aus `Skript-Module`
2. Setzen des Word-Templatepfads auf den Desktop des aktuellen Benutzers
3. Anheften des Desktops an den Schnellzugriff
4. Start des Transcript-Loggings in `4. Logs`
5. Prüfung, ob Word und Excel per COM verfügbar sind
6. Automatische Bereitstellung der neuesten Nuera-Dateien
7. Office- und Explorer-Konfiguration per Registry und optional per COM
8. Kopieren von Vorlagen und Schnellzugriff-Dateien
9. Erzeugung der Kandidatenordner aus Excel oder CSV
10. Anwenden von Taskleisten- und Proxy-Einstellungen

## Parameter

| Parameter | Typ | Standard | Beschreibung |
| --- | --- | --- | --- |
| `-Proxy` | `On`, `Off`, `Skip` | `Skip` | Aktiviert, deaktiviert oder überspringt die Proxy-Konfiguration |
| `-ProxyServer` | `String` | `192.168.0.1:8080` | Serveradresse für den Proxy |
| `-ProxyBypass` | `String` | Microsoft-/Office-Ausnahmeliste | Ausnahmen für die Proxy-Registry |
| `-ExcelListPath` | `String` | automatisch `1. Anpassen\AP1-TN.xlsx` | Quelle der Teilnehmerliste |
| `-CsvFallbackPath` | `String` | leer | CSV-Alternative für den Fall, dass Excel/COM nicht genutzt werden kann |
| `-MaxRows` | `Int` | `500` | Maximal zu lesende Datensätze |
| `-Quiet` | `Switch` | aus | Unterdrückt die interaktive Proxy-Bestätigung |
| `-RegistryOnly` | `Switch` | aus | Erzwingt reine Registry-Ausführung ohne COM |

## Module und Zuständigkeiten

| Modul | Verantwortung |
| --- | --- |
| `AP1-Logging.psm1` | Transcript-Logging, Konsolenausgabe, Rotation alter Logs |
| `AP1-Office.psm1` | First-Run-Erkennung und Office-Initialisierung |
| `AP1-Registry.psm1` | Registry-Helfer, COM-Bereinigung, Verfügbarkeitsprüfungen |
| `AP1-Folders.psm1` | Kandidatenordner aus Excel/CSV und Desktop-Bereitstellung |
| `AP1-Templates.psm1` | Vorlagen und Office-Schnellzugriff |
| `AP1-System.psm1` | Proxy, Taskleiste, ZIP-Entpackung, Prozesssteuerung |
| `AP1-Nuera.psm1` | Download, Entpacken und Kopieren der Nuera-Dateien |
| `AP1-QuickAccess.psm1` | Schnellzugriff und Desktop-Shortcuts |
| `AP1-Utils.psm1` | Hilfsfunktionen für Text und Verknüpfungen |

## Besondere Laufzeitlogik

### COM-Fallback

- Standardmäßig versucht das Skript, Word und Excel per COM zu initialisieren.
- Falls einer der beiden Starts fehlschlägt, wird automatisch in den Registry-Fallback gewechselt.
- Mit `-RegistryOnly` kann dieser Modus direkt erzwungen werden.

### Teilnehmerordner

- Primärquelle ist `1. Anpassen\AP1-TN.xlsx`.
- Falls COM/Excel nicht genutzt werden kann, wird bei gesetztem `-CsvFallbackPath` automatisch die CSV-Verarbeitung genutzt.
- Temporäre Ordner unter `2. Bei Bedarf anpassen\Ordner` werden nach dem Kopieren auf den Desktop wieder aufgeräumt.

### Nuera-Dateien

- Der aktuelle Stand lädt die erste verfügbare Datei aus einer priorisierten Liste (`nuera2026_f.zip`, `nuera2026_h.zip`, ...).
- Nach dem Download wird das Archiv entpackt und der extrahierte Ordner auf den Desktop kopiert.

## Logging

- Logdateien werden als `AP1_Prep_yyyyMMdd_HHmmss.log` unter `4. Logs` abgelegt.
- Es bleiben maximal fünf Logdateien erhalten; ältere Logs werden automatisch gelöscht.
- `Write-SafeOutput` wird für sichtbare Statusmeldungen genutzt.

## Prüfergebnis für v1.0.10

Für diese Version wurden insbesondere folgende Probleme bereinigt:

- `Write-SafeOutput` akzeptiert nun wieder `-ForegroundColor` und gibt Meldungen tatsächlich aus.
- Die dokumentierten Parameter `-Quiet` und `-RegistryOnly` sind nun im Hauptskript nutzbar.
- Die Hauptfunktion verarbeitet Parameter jetzt explizit und zuverlässig.
- Doppelte Initialisierungsblöcke für Modulimport und Encoding wurden entfernt.
- Irreführende COM-Statusmeldung und doppelte Fehlerausgaben wurden bereinigt.

## Troubleshooting

| Problem | Ursache | Lösung |
| --- | --- | --- |
| Keine sichtbare Start-/Endmeldung | veraltete `Write-SafeOutput`-Implementierung | ab `v1.0.10` behoben |
| Proxy-Rückfrage stört Automatisierung | interaktive Bestätigung aktiv | `-Quiet` setzen |
| Word/Excel COM startet nicht | First-Run, Lizenzdialoge oder Office-Profilproblem | Office einmal manuell starten oder `-RegistryOnly` verwenden |
| Excel-Datei fehlt | `AP1-TN.xlsx` nicht vorhanden oder anderer Pfad nötig | `-ExcelListPath` explizit angeben |
| CSV statt Excel erforderlich | COM/Excel nicht nutzbar | `-CsvFallbackPath` mitgeben |

## Relevante Dateien

- Hauptskript: `AP1-Konfigurator.ps1`
- Start-Batch: `AP1-Konfigurator.bat`
- Kurzinfo: `docs/Liesmich.txt`
- Änderungsverlauf: `CHANGELOG.md`
- Release-Hinweise: `RELEASE_NOTES_v1.0.10.md`
