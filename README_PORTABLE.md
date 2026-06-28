# AP1-Konfigurator-Portable v1.0.11

Dies ist die **EXE-Variante** des Projekts `AP1-Konfigurator`.

## Inhalt des Releases

Dieses Paket enthält bewusst nur:

- `AP1-Konfigurator-Portable.exe`
- `data/`
- `docs/`
- `README.md`

## Start

1. `AP1-Konfigurator-Portable.exe` starten.
2. Beim Start werden die eingebetteten Laufzeitdateien nach `%LOCALAPPDATA%\AP1-Konfigurator-Portable\v1.0.11` kopiert.
3. Die EXE verwendet anschließend diese lokale Arbeitskopie.
4. `data/` und `docs/` aus diesem Release werden ebenfalls dorthin synchronisiert.

## Hinweise

- Die technischen PowerShell-Startdateien und `Skript-Module` sind **in der EXE eingebettet** und liegen nicht separat im Release.
- Änderungen an `data/` im Release-Verzeichnis werden beim nächsten Start erneut in die lokale Arbeitskopie übernommen.
- Weiterführende Informationen liegen unter `docs/`.
