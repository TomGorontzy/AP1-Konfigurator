# RELEASE-PROZESS

## Ziel

Für `AP1-Konfigurator` sollen Releases **grundsätzlich als EXE-Variante** veröffentlicht werden.

Primäres Endanwender-Artefakt:

- `release/AP1-Konfigurator-Portable-vX.Y.Z.zip`

## Versionierung

Die Build-Pipeline liest die Version aus:

- `src/build_info.py`

Beispiel:

- `1.0.14` → Release-Tag `v1.0.14`

## Release-Checkliste

1. Dokumentation prüfen:
   - `README.md`
   - `docs/DOKUMENTATION_ANWENDER.md`
   - `docs/DOKUMENTATION_TECHNIK.md`
   - `docs/KURZDOKUMENTATION.txt`
   - `docs/DOKUMENTATION_CHECKLISTE.md`
2. Release-Hinweise prüfen oder anlegen:
   - `release/RELEASE_NOTES_vX.Y.Z.md`
   - Für neue Dateien die Vorlage `release/RELEASE_NOTES_TEMPLATE.md` verwenden.
   - Inhaltlich an der Struktur von `release/RELEASE_NOTES_v1.0.10.md` orientieren:
     - `## Highlights`
     - `## Änderungen im Detail`
     - `## Validierung`
     - `## Artefakte` / `## Artefakte und GitHub-Hinweis`
   - Änderungen nicht nur als knappe Stichworte notieren, sondern nachvollziehbar in Funktionsblöcke und Dokumentation gliedern.
3. Python-Buildumgebung einrichten:
   - `./src/setup.ps1`
4. EXE-Build ausführen:
   - `./src/build.ps1`
   - oder ohne Versionssprung: `./src/build.ps1 -NoVersionBump`
5. Artefakte validieren:
   - `dist/AP1-Konfigurator-Portable-vX.Y.Z/AP1-Konfigurator-Portable.exe`
   - `release/AP1-Konfigurator-Portable-vX.Y.Z/`
   - `release/AP1-Konfigurator-Portable-vX.Y.Z.zip`
   - Im Release-Ordner befinden sich nur `AP1-Konfigurator-Portable.exe`, `data/`, `docs/` und `README.md`
   - Die ZIP-Datei ist flach und enthält keinen zusätzlichen Oberordner
6. Veröffentlichung ausführen:
   - `./src/publish_release.ps1 -Version vX.Y.Z`
7. GitHub prüfen:
   - Tag sichtbar
   - Release vorhanden
   - EXE-ZIP-Asset vorhanden

## Standardablauf lokal

### 1) Buildumgebung vorbereiten

```powershell
./src/setup.ps1
```

### 2) EXE bauen

```powershell
./src/build.ps1
```

### 3) Release veröffentlichen

```powershell
./src/publish_release.ps1 -Version vX.Y.Z
```

## GitHub Actions

Der Workflow `.github/workflows/release.yml` baut bei Tag-Pushes `v*` automatisch die EXE-Variante und lädt das resultierende ZIP in den GitHub-Release hoch.

## Hinweise

- Der aktuelle EXE-Launcher enthält die PowerShell-Startskripte und Module eingebettet.
- Beim Start werden die eingebetteten Laufzeitdateien nach `%LOCALAPPDATA%\AP1-Konfigurator-Portable\vX.Y.Z` kopiert; zusätzlich wird `%LOCALAPPDATA%\AP1-Konfigurator-Portable\current` als aktuelle Arbeitskopie aktualisiert.
- Ältere `%LOCALAPPDATA%\AP1-Konfigurator-Portable\v*`-Ordner werden beim Start automatisch bereinigt.
- Die EXE dient als bevorzugte Endanwender-Startvariante.
- Das Release-Paket selbst bleibt schlank und enthält nur `AP1-Konfigurator-Portable.exe`, `data/`, `docs/` und `README.md`.
- Das ZIP ist flach aufgebaut und hat keinen zusätzlichen Oberordner.
