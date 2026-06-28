from __future__ import annotations

import shutil
import sys
from pathlib import Path

from build_info import BUILD_INFO

ROOT = Path(__file__).resolve().parents[1]
VERSION = f"v{BUILD_INFO['version']}"
ARTIFACT_NAME = BUILD_INFO['artifact_name']
DIST_ROOT = ROOT / 'dist'
RELEASE_ROOT = ROOT / 'release'
PACKAGE_DIR = DIST_ROOT / f'{ARTIFACT_NAME}-{VERSION}'
RELEASE_DIR = RELEASE_ROOT / f'{ARTIFACT_NAME}-{VERSION}'
EXE_SOURCE = DIST_ROOT / f'{ARTIFACT_NAME}.exe'

COPY_MAP = {
    ROOT / 'AP1-Konfigurator.ps1': 'data/AP1-Konfigurator.ps1',
    ROOT / 'AP1-Konfigurator.bat': 'data/AP1-Konfigurator.bat',
    ROOT / 'Proxy-Deaktivieren.bat': 'data/Proxy-Deaktivieren.bat',
    ROOT / 'README.md': 'README.md',
    ROOT / 'CHANGELOG.md': 'data/CHANGELOG.md',
    ROOT / 'DOCUMENTATION.md': 'data/DOCUMENTATION.md',
    ROOT / f'RELEASE_NOTES_{VERSION}.md': f'RELEASE_NOTES_{VERSION}.md',
    ROOT / 'Skript-Module': 'data/Skript-Module',
    ROOT / '1. Anpassen': 'data/1. Anpassen',
    ROOT / '2. Bei Bedarf anpassen': 'data/2. Bei Bedarf anpassen',
    ROOT / '3. Nuera-Dateien': 'data/3. Nuera-Dateien',
    ROOT / 'docs': 'docs',
}


def reset_dir(path: Path) -> None:
    if path.exists():
        shutil.rmtree(path)
    path.mkdir(parents=True, exist_ok=True)


def try_reset_dir(path: Path) -> bool:
    try:
        reset_dir(path)
        return True
    except PermissionError as exc:
        print(f'warnung: {path} konnte nicht bereinigt werden ({exc}). Spiegelordner wird übersprungen.')
        return False


def copy_item(source: Path, relative_target: str, package_dir: Path) -> None:
    target = package_dir / relative_target
    target.parent.mkdir(parents=True, exist_ok=True)
    if source.is_dir():
        shutil.copytree(source, target, dirs_exist_ok=True)
    else:
        shutil.copy2(source, target)


def main() -> int:
    if not EXE_SOURCE.exists():
        raise FileNotFoundError(f'PyInstaller-Ausgabe fehlt: {EXE_SOURCE}')

    reset_dir(PACKAGE_DIR)
    release_dir_ready = try_reset_dir(RELEASE_DIR)

    shutil.copy2(EXE_SOURCE, PACKAGE_DIR / f'{ARTIFACT_NAME}.exe')
    for source, target in COPY_MAP.items():
        if source.exists():
            copy_item(source, target, PACKAGE_DIR)

    (PACKAGE_DIR / 'data' / '4. Logs').mkdir(parents=True, exist_ok=True)
    if release_dir_ready:
        shutil.copytree(PACKAGE_DIR, RELEASE_DIR, dirs_exist_ok=True)

    archive_base = RELEASE_ROOT / f'{ARTIFACT_NAME}-{VERSION}'
    zip_path = shutil.make_archive(str(archive_base), 'zip', root_dir=PACKAGE_DIR.parent, base_dir=PACKAGE_DIR.name)
    print(f'paketiert: {zip_path}')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
