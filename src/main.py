from __future__ import annotations

import ctypes
import importlib.util
import os
import shutil
import subprocess
import sys
from pathlib import Path

CREATE_NEW_CONSOLE = 0x00000010
MB_ICONERROR = 0x00000010
MB_OK = 0x00000000
EMBEDDED_FILES = (
    'AP1-Konfigurator.ps1',
    'AP1-Konfigurator.bat',
    'Proxy-Deaktivieren.bat',
)
EMBEDDED_DIRS = (
    'Skript-Module',
    'data',
    'docs',
)
PACKAGE_CONTENT_DIRS = (
    'data',
    'docs',
)


def show_error(message: str, title: str = 'AP1-Konfigurator-Portable') -> None:
    try:
        ctypes.windll.user32.MessageBoxW(None, message, title, MB_OK | MB_ICONERROR)
    except Exception:
        sys.stderr.write(f'{title}: {message}\n')


def resolve_runtime_dir() -> Path:
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parents[1]


def resolve_source_dir() -> Path:
    return Path(__file__).resolve().parent


def load_build_info() -> dict[str, str]:
    candidates = []
    if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
        candidates.append(Path(getattr(sys, '_MEIPASS')) / 'build_info.py')
    candidates.append(resolve_source_dir() / 'build_info.py')
    candidates.append(resolve_runtime_dir() / 'src' / 'build_info.py')

    for candidate in candidates:
        if not candidate.exists():
            continue
        spec = importlib.util.spec_from_file_location('ap1_build_info', candidate)
        if spec and spec.loader:
            module = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(module)
            return module.BUILD_INFO

    raise ModuleNotFoundError('build_info.py konnte weder als Datei noch im Bundle geladen werden.')


BUILD_INFO = load_build_info()
ARTIFACT_NAME = BUILD_INFO['artifact_name']
VERSION = f"v{BUILD_INFO['version']}"


def resolve_bundle_dir() -> Path:
    if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
        return Path(getattr(sys, '_MEIPASS')).resolve()
    return Path(__file__).resolve().parents[1]


def resolve_local_appdata_root() -> Path:
    local_appdata = os.environ.get('LOCALAPPDATA')
    if not local_appdata:
        local_appdata = str(Path.home() / 'AppData' / 'Local')
    return Path(local_appdata) / ARTIFACT_NAME


def resolve_local_appdata_dir() -> Path:
    return resolve_local_appdata_root() / VERSION


def resolve_current_app_dir() -> Path:
    return resolve_local_appdata_root() / 'current'


def copy_file(source: Path, target: Path) -> None:
    target.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(source, target)


def sync_directory(source: Path, target: Path) -> None:
    if not source.exists():
        return
    target.mkdir(parents=True, exist_ok=True)
    for item in source.iterdir():
        destination = target / item.name
        if item.is_dir():
            sync_directory(item, destination)
        else:
            copy_file(item, destination)


def sync_embedded_runtime(bundle_dir: Path, app_dir: Path) -> None:
    for file_name in EMBEDDED_FILES:
        source = bundle_dir / file_name
        if source.exists():
            copy_file(source, app_dir / file_name)

    for dir_name in EMBEDDED_DIRS:
        source = bundle_dir / dir_name
        if source.exists():
            sync_directory(source, app_dir / dir_name)


def sync_release_content(runtime_dir: Path, app_dir: Path) -> None:
    for dir_name in PACKAGE_CONTENT_DIRS:
        source = runtime_dir / dir_name
        if source.exists():
            sync_directory(source, app_dir / dir_name)


def reset_dir(path: Path) -> None:
    if path.exists():
        shutil.rmtree(path)
    path.mkdir(parents=True, exist_ok=True)


def prune_old_version_dirs(root_dir: Path, current_version: str) -> None:
    if not root_dir.exists():
        return
    for child in root_dir.iterdir():
        if not child.is_dir():
            continue
        if child.name == 'current' or child.name == current_version:
            continue
        if not child.name.startswith('v'):
            continue
        try:
            shutil.rmtree(child)
        except PermissionError:
            continue


def prepare_current_alias(source_dir: Path) -> Path:
    current_dir = resolve_current_app_dir()
    try:
        reset_dir(current_dir)
        sync_directory(source_dir, current_dir)
        return current_dir
    except PermissionError:
        return source_dir


def prepare_app_dir(runtime_dir: Path) -> Path:
    bundle_dir = resolve_bundle_dir()
    root_dir = resolve_local_appdata_root()
    app_dir = resolve_local_appdata_dir()
    app_dir.mkdir(parents=True, exist_ok=True)
    sync_embedded_runtime(bundle_dir, app_dir)
    sync_release_content(runtime_dir, app_dir)
    (app_dir / 'data' / '4. Logs').mkdir(parents=True, exist_ok=True)
    prune_old_version_dirs(root_dir, VERSION)
    return prepare_current_alias(app_dir)


def find_launcher(runtime_dir: Path) -> Path:
    candidates = [
        runtime_dir / 'AP1-Konfigurator.bat',
        runtime_dir / 'AP1-Konfigurator.ps1',
        runtime_dir / 'data' / 'AP1-Konfigurator.bat',
        runtime_dir / 'data' / 'AP1-Konfigurator.ps1',
    ]
    for candidate in candidates:
        if candidate.exists():
            return candidate
    raise FileNotFoundError('Kein Startskript gefunden. Erwartet wurden AP1-Konfigurator.bat oder AP1-Konfigurator.ps1.')


def build_command(launcher: Path) -> list[str]:
    args = sys.argv[1:]
    suffix = launcher.suffix.lower()
    if suffix == '.bat':
        return ['cmd.exe', '/c', str(launcher), *args]
    if suffix == '.ps1':
        return [
            'powershell.exe',
            '-NoProfile',
            '-ExecutionPolicy',
            'Bypass',
            '-File',
            str(launcher),
            *args,
        ]
    raise RuntimeError(f'Nicht unterstütztes Startskript: {launcher.name}')


def main() -> int:
    runtime_dir = resolve_runtime_dir()
    try:
        app_dir = prepare_app_dir(runtime_dir)
        launcher = find_launcher(app_dir)
        command = build_command(launcher)
        subprocess.Popen(command, cwd=str(app_dir), creationflags=CREATE_NEW_CONSOLE)
        return 0
    except Exception as exc:
        show_error(str(exc))
        return 1


if __name__ == '__main__':
    raise SystemExit(main())
