
# AP1-Konfigurator

Automatisiert die Einrichtung von Prüfungsrechnern für die Abschlussprüfung Teil 1 (AP1).

## Features

- Office-Konfiguration (Word & Excel, Formatierung, Vorlagen)
- Kandidatenordner-Erstellung aus Excel/CSV
- Registry- und COM-basierte Einstellungen
- Logging im Ordner `4. Logs`
- Robustes Fallback-System (Registry-only bei COM-Fehlern)

## Schnellstart

```powershell
.\AP1-Konfigurator.ps1
```

Parameter:
- `-Nuera` (Nuera-Dateien laden)
- `-Proxy On|Off|Skip` (Proxy-Konfiguration)
- `-ExcelListPath` (Pfad zur Teilnehmerliste)
- `-CsvFallbackPath` (CSV-Fallback)

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

## Dokumentation

- Technische Details: [DOCUMENTATION.md](./DOCUMENTATION.md)
- Inline-Kommentare im Skript

## Support

Bei Fragen oder Problemen: Siehe Dokumentation oder wende dich an die Projektverantwortlichen.

---

*Letzte Aktualisierung: 21. Februar 2026*