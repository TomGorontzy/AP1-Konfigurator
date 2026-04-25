
# AP1-Konfigurator – Vollständige Dokumentation

## Übersicht

Der AP1-Konfigurator automatisiert die Einrichtung von Prüfungsrechnern für die Abschlussprüfung Teil 1 (AP1). Das PowerShell-Skript sorgt für eine standardisierte Office-Umgebung, legt Kandidatenordner an und protokolliert alle Schritte.

---

## Installation & Ausführung

1. PowerShell mit Administratorrechten öffnen (optional, meist nicht nötig)
2. Skript ausführen:
   ```powershell
   .\AP1-Konfigurator.ps1
   ```
3. Parameter nach Bedarf verwenden:
   - `-Nuera` (lädt aktuelle Nuera-Dateien)
   - `-Proxy On|Off|Skip` (Proxy-Konfiguration)
   - `-ExcelListPath` (Pfad zur Teilnehmerliste)
   - `-CsvFallbackPath` (CSV-Fallback)

---

## Parameter

| Name             | Typ    | Standard                  | Beschreibung                       |
|------------------|--------|---------------------------|------------------------------------|
| Proxy            | String | Skip                      | Proxy-Modus (On/Off/Skip)          |
| ProxyServer      | String | 192.168.0.1:8080          | Proxy-Server-Adresse               |
| ProxyBypass      | String | domain.local              | Proxy-Bypass-Liste                 |
| Nuera            | Switch | -                         | Nuera-Dateien laden                |
| ExcelListPath    | String | ./1. Anpassen/AP1-TN.xlsx | Teilnehmerliste                    |
| CsvFallbackPath  | String | -                         | CSV-Fallback bei Excel-Fehlern     |
| MaxRows          | Int    | 500                       | Max. Teilnehmerzeilen              |
| Quiet            | Switch | -                         | Reduzierte Ausgabe                 |
| RegistryOnly     | Switch | -                         | Erzwingt Registry-Modus            |

---

## Funktionsweise

### Office-Konfiguration
- Initialisiert Word und Excel (First-Run-Erkennung)
- Setzt Standardschriftart (Arial 11pt), Zeilenabstand (1,1x), Absatzabstand (0pt)
- Deaktiviert unerwünschte Autokorrektur- und Formatierungsoptionen
- Kopiert und ersetzt Vorlagen (Normal.dotm, Mappe.xltx)

### Kandidatenordner
- Liest Teilnehmer aus Excel oder CSV
- Erstellt Ordnerstruktur für jeden Kandidaten
- Kopiert Ordner auf Desktop

### Logging
- Erstellt Logdateien im Ordner `4. Logs`
- Maximal 5 Logs werden behalten, ältere werden automatisch gelöscht
- Vollständige Protokollierung aller Aktionen

### Fallback-Mechanismen
- Automatische Umschaltung auf Registry-only-Modus bei COM-Fehlern
- CSV-Fallback bei Excel-Problemen

---

## Projektstruktur

```
AP1-Konfigurator/
├── AP1-Konfigurator.ps1      # Hauptskript
├── Skript-Module/            # Alle PowerShell-Module
├── 1. Anpassen/              # Teilnehmerliste (Excel)
├── 2. Bei Bedarf anpassen/   # Vorlagen, Symbolleisten
├── 3. Nuera-Dateien/         # Prüfungsunterlagen
├── 4. Logs/                  # Ausführungsprotokolle
├── _Archive/                 # Entwicklungsversionen
```

---

## Troubleshooting

| Problem                        | Lösung                                         |
|--------------------------------|------------------------------------------------|
| COM-Start fehlgeschlagen       | Registry-only-Modus wird automatisch aktiviert |
| Font bleibt Aptos              | Font-Substitution + Template-Methode           |
| Zeilenabstand falsch           | Selection-basierte Formatierung                |
| Excel-Liste nicht lesbar       | CSV-Fallback aktivieren                        |
| Log nicht im richtigen Ordner  | Logging-Funktion im Modul prüfen               |

---

## Erweiterung & Anpassung

- Neue Module einfach in `Skript-Module` ablegen
- Logging und Konfiguration über Parameter steuerbar
- Anpassungen an Vorlagen und Registry-Einstellungen im jeweiligen Modul

---

## Kontakt & Support

- Hauptdokumentation: `README.md`
- Technische Details: `DOCUMENTATION.md`
- Inline-Kommentare im Skript

---

*Letzte Aktualisierung: 21. Februar 2026*