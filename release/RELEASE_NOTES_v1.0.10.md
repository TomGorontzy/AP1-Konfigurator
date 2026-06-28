# AP1-Konfigurator v1.0.10

## Highlights

- Die sichtbaren Statusmeldungen des Hauptskripts funktionieren wieder zuverlässig.
- Dokumentierte Schalter `-Quiet` und `-RegistryOnly` sind nun technisch korrekt verdrahtet.
- README, technische Dokumentation und Kurzinfo wurden auf den aktuellen Funktionsstand gebracht.

## Änderungen im Detail

### Skriptstabilität

- `Skript-Module/AP1-Logging.psm1`
  - `Write-SafeOutput` gibt wieder Ausgaben auf der Konsole aus.
  - Unterstützung für `-ForegroundColor` ergänzt.
- `AP1-Konfigurator.ps1`
  - Parametertabelle im Skript auf den tatsächlich verwendeten Satz erweitert.
  - Standardwerte für Proxy-Server und Proxy-Bypass vereinheitlicht.
  - doppelte Initialisierungsblöcke entfernt.
  - COM-Statusmeldung korrigiert.
  - doppelte Fehlerausgabe beim Excel-Pfad entfernt.
  - Hauptaufruf auf sauberes Splatting umgestellt.

### Dokumentation

- `README.md` vollständig auf den Stand `v1.0.10` aktualisiert.
- `DOCUMENTATION.md` technisch überarbeitet und um Prüfergebnis ergänzt.
- `docs/Liesmich.txt` von Platzhalter auf nutzbare Kurzinfo umgestellt.
- `CHANGELOG.md` neu angelegt.

## Validierung

- Syntaxprüfung der angepassten PowerShell-Dateien ohne gemeldete Fehler.
- Workspace-Fehlerprüfung für die geänderten Dateien ohne Befund.

## Artefakte und GitHub-Hinweis

- Diese Release-Notiz ist für den GitHub-Release `v1.0.10` vorbereitet.
- Falls zusätzlich ein Binär-/Portable-Artefakt veröffentlicht werden soll, muss es aus der tatsächlich gepflegten Build-Pipeline erzeugt und anschließend an den GitHub-Release angehängt werden.
