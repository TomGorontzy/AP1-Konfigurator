# AP1-Konfigurator-Portable v1.0.11

Dies ist die **EXE-Variante** des Projekts `AP1-Konfigurator`.

## Inhalt des Releases

Dieses Paket enthält bewusst nur:

- `AP1-Konfigurator-Portable.exe`
- `data/`
- `docs/`
- `README.md`

Die ZIP-Datei ist **flach** aufgebaut und enthält keinen zusätzlichen Oberordner.

## Start

1. `AP1-Konfigurator-Portable.exe` starten.
2. Beim Start werden die eingebetteten Laufzeitdateien nach `%LOCALAPPDATA%\AP1-Konfigurator-Portable\v1.0.11` kopiert.
3. Zusätzlich wird `%LOCALAPPDATA%\AP1-Konfigurator-Portable\current` als aktuelle Arbeitskopie aktualisiert.
4. Die EXE verwendet anschließend bevorzugt diese `current`-Arbeitskopie.
5. `data/` und `docs/` aus diesem Release werden ebenfalls dorthin synchronisiert.
6. Ältere `%LOCALAPPDATA%\AP1-Konfigurator-Portable\v*`-Ordner werden dabei automatisch bereinigt.

## Hinweise

- Die technischen PowerShell-Startdateien und `Skript-Module` sind **in der EXE eingebettet** und liegen nicht separat im Release.
- Änderungen an `data/` im Release-Verzeichnis werden beim nächsten Start erneut in die lokale Arbeitskopie übernommen.
- Für Bereinigung oder Diagnose können sowohl der versionsbezogene Ordner als auch `%LOCALAPPDATA%\AP1-Konfigurator-Portable\current` geprüft werden.
- Ältere versionierte Arbeitsordner werden beim Start automatisch entfernt, sofern sie nicht mehr der aktuellen Version entsprechen.
- Das Release-ZIP beginnt direkt mit `AP1-Konfigurator-Portable.exe`, `data/`, `docs/` und `README.md`.
- Weiterführende Informationen liegen unter `docs/`.
