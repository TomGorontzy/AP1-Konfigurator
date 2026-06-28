# RELEASE-PROZESS

## Ziel

Für `AP1-Konfigurator` sollen Releases **grundsätzlich als EXE-Variante** veröffentlicht werden.

Primäres Endanwender-Artefakt:

- `release/AP1-Konfigurator-Portable-vX.Y.Z.zip`

Ergänzendes Artefakt:

- `release/AP1-Konfigurator-vX.Y.Z.zip` (Skript-/Quellpaket, optional)

## Versionierung

Die Build-Pipeline liest die Version aus:

- `src/build_info.py`

Beispiel:

- `1.0.11` → Release-Tag `v1.0.11`

## Release-Checkliste

1. Dokumentation prüfen:
   - `README.md`
   - `docs/DOKUMENTATION_ANWENDER.md`
   - `docs/DOKUMENTATION_TECHNIK.md`
   - `docs/KURZDOKUMENTATION.txt`
   - `docs/DOKUMENTATION_CHECKLISTE.md`
2. Release-Hinweise prüfen oder anlegen:
   - `RELEASE_NOTES_vX.Y.Z.md`
3. Python-Buildumgebung einrichten:
   - `./setup.ps1`
4. EXE-Build ausführen:
   - `./build.ps1`
   - oder ohne Versionssprung: `./build.ps1 -NoVersionBump`
5. Artefakte validieren:
   - `dist/AP1-Konfigurator-Portable-vX.Y.Z/AP1-Konfigurator-Portable.exe`
   - `release/AP1-Konfigurator-Portable-vX.Y.Z/`
   - `release/AP1-Konfigurator-Portable-vX.Y.Z.zip`
6. Veröffentlichung ausführen:
   - `./publish_release.ps1 -Version vX.Y.Z`
7. GitHub prüfen:
   - Tag sichtbar
   - Release vorhanden
   - EXE-ZIP-Asset vorhanden

## Standardablauf lokal

### 1) Buildumgebung vorbereiten

```powershell
./setup.ps1
```

### 2) EXE bauen

```powershell
./build.ps1
```

### 3) Release veröffentlichen

```powershell
./publish_release.ps1 -Version vX.Y.Z
```

## GitHub Actions

Der Workflow `.github/workflows/release.yml` baut bei Tag-Pushes `v*` automatisch die EXE-Variante und lädt das resultierende ZIP in den GitHub-Release hoch.

## Hinweise

- Der aktuelle EXE-Launcher ist bewusst minimal und startet das mitgelieferte Batch-/PowerShell-Skript.
- Die EXE dient als bevorzugte Endanwender-Startvariante.
- Für Notfälle kann zusätzlich ein Skriptpaket veröffentlicht werden, dieses ist jedoch **nicht** das primäre Release-Format.
