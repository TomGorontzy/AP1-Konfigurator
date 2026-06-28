# Changelog

## v1.0.10 - 2026-06-28

### Behoben

- `Write-SafeOutput` unterstützt wieder `-ForegroundColor` und gibt Statusmeldungen sichtbar aus.
- Die dokumentierten Schalter `-Quiet` und `-RegistryOnly` sind nun im Hauptskript korrekt nutzbar.
- Die Hauptfunktion `Start-AP1Konfiguration` verarbeitet Parameter jetzt explizit statt indirekt und fehleranfällig.
- Die irreführende Meldung zum Registry-only-Modus bei erfolgreicher COM-Erkennung wurde korrigiert.
- Doppelte Fehlerausgabe beim fehlenden Excel-Pfad wurde entfernt.

### Bereinigt

- Doppelte Initialisierung für Encoding und Modulimporte im Hauptskript entfernt.
- README, technische Dokumentation und Kurzinfo auf den tatsächlichen Funktionsumfang aktualisiert.
- Release-Dokumentation für `v1.0.10` ergänzt.

## v1.0.9

- Vorheriger veröffentlichter Stand im Repository.

## v1.0.8

- Vorheriger veröffentlichter Stand im Repository.

## v1.0.7

- Vorheriger veröffentlichter Stand im Repository; siehe `release/RELEASE_NOTES_v1.0.7.md`.
