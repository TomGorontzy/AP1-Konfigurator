from __future__ import annotations

import shutil
import sys
import tempfile
from pathlib import Path

from build_info import BUILD_INFO

ROOT = Path(__file__).resolve().parents[1]
VERSION = f"v{BUILD_INFO['version']}"
ARTIFACT_NAME = BUILD_INFO['artifact_name']
LEGACY_ARTIFACT_NAME = 'AP1-Konfigurator'
DIST_ROOT = ROOT / 'dist'
RELEASE_ROOT = ROOT / 'release'
PACKAGE_DIR = DIST_ROOT / f'{ARTIFACT_NAME}-{VERSION}'
RELEASE_DIR = RELEASE_ROOT / f'{ARTIFACT_NAME}-{VERSION}'
EXE_SOURCE = DIST_ROOT / f'{ARTIFACT_NAME}.exe'

COPY_MAP = {
    ROOT / 'README_PORTABLE.md': 'README.md',
    ROOT / '1. Anpassen': 'data/1. Anpassen',
    ROOT / '2. Bei Bedarf anpassen': 'data/2. Bei Bedarf anpassen',
    ROOT / '3. Nuera-Dateien': 'data/3. Nuera-Dateien',
    ROOT / 'docs': 'docs',
}

OBSOLETE_PACKAGE_PATHS = (
    Path('AP1-Konfigurator.ps1'),
    Path('AP1-Konfigurator.bat'),
    Path('Proxy-Deaktivieren.bat'),
    Path(f'RELEASE_NOTES_{VERSION}.md'),
    Path('Skript-Module'),
    Path('data/AP1-Konfigurator.bat'),
    Path('data/AP1-Konfigurator.ps1'),
    Path('data/Proxy-Deaktivieren.bat'),
    Path('data/CHANGELOG.md'),
    Path('data/DOCUMENTATION.md'),
    Path('data/Skript-Module'),
)


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


def try_remove_path(path: Path) -> None:
    try:
        if path.is_dir():
            shutil.rmtree(path)
        elif path.exists():
            path.unlink()
    except PermissionError as exc:
        print(f'warnung: veraltetes Artefakt {path} konnte nicht entfernt werden ({exc}).')


def prune_obsolete_paths(package_root: Path) -> None:
    for relative_path in OBSOLETE_PACKAGE_PATHS:
        try_remove_path(package_root / relative_path)


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

    legacy_release_dir = RELEASE_ROOT / f'{LEGACY_ARTIFACT_NAME}-{VERSION}'
    legacy_release_zip = RELEASE_ROOT / f'{LEGACY_ARTIFACT_NAME}-{VERSION}.zip'
    try_remove_path(legacy_release_dir)
    try_remove_path(legacy_release_zip)

    with tempfile.TemporaryDirectory(prefix='ap1-konfigurator-') as temp_dir:
        staging_root = Path(temp_dir)
        staging_package_dir = staging_root / f'{ARTIFACT_NAME}-{VERSION}'
        staging_package_dir.mkdir(parents=True, exist_ok=True)

        dist_dir_ready = try_reset_dir(PACKAGE_DIR)
        release_dir_ready = try_reset_dir(RELEASE_DIR)

        shutil.copy2(EXE_SOURCE, staging_package_dir / f'{ARTIFACT_NAME}.exe')
        for source, target in COPY_MAP.items():
            if source.exists():
                copy_item(source, target, staging_package_dir)

        (staging_package_dir / 'data' / '4. Logs').mkdir(parents=True, exist_ok=True)

        if dist_dir_ready:
            shutil.copytree(staging_package_dir, PACKAGE_DIR, dirs_exist_ok=True)
        else:
            print(f'warnung: {PACKAGE_DIR} konnte nicht aktualisiert werden. Dist-Spiegelordner wird übersprungen.')
        prune_obsolete_paths(PACKAGE_DIR)

        if release_dir_ready:
            shutil.copytree(staging_package_dir, RELEASE_DIR, dirs_exist_ok=True)
        prune_obsolete_paths(RELEASE_DIR)

        archive_base = RELEASE_ROOT / f'{ARTIFACT_NAME}-{VERSION}'
        zip_path = shutil.make_archive(str(archive_base), 'zip', root_dir=staging_package_dir.parent, base_dir=staging_package_dir.name)

    print(f'paketiert: {zip_path}')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
