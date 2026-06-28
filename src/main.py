from __future__ import annotations

import ctypes
import os
import subprocess
import sys
from pathlib import Path

CREATE_NEW_CONSOLE = 0x00000010
MB_ICONERROR = 0x00000010
MB_OK = 0x00000000


def show_error(message: str, title: str = 'AP1-Konfigurator-Portable') -> None:
    try:
        ctypes.windll.user32.MessageBoxW(None, message, title, MB_OK | MB_ICONERROR)
    except Exception:
        sys.stderr.write(f'{title}: {message}\n')


def resolve_runtime_dir() -> Path:
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parents[1]


def find_launcher(runtime_dir: Path) -> Path:
    candidates = [
        runtime_dir / 'data' / 'AP1-Konfigurator.bat',
        runtime_dir / 'AP1-Konfigurator.bat',
        runtime_dir / 'data' / 'AP1-Konfigurator.ps1',
        runtime_dir / 'AP1-Konfigurator.ps1',
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
        launcher = find_launcher(runtime_dir)
        command = build_command(launcher)
        subprocess.Popen(command, cwd=str(launcher.parent), creationflags=CREATE_NEW_CONSOLE)
        return 0
    except Exception as exc:
        show_error(str(exc))
        return 1


if __name__ == '__main__':
    raise SystemExit(main())
