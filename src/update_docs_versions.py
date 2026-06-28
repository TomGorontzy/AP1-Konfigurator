from __future__ import annotations

import re
from pathlib import Path

from build_info import BUILD_INFO

ROOT = Path(__file__).resolve().parents[1]
VERSION = f"v{BUILD_INFO['version']}"
DATE_TEXT = '28. Juni 2026'

REPLACEMENTS: dict[Path, list[tuple[str, str]]] = {
    ROOT / 'README.md': [
        (r"Aktueller Stand: \*\*v\d+\.\d+\.\d+\*\* · Letzte Aktualisierung: \*\*[^*]+\*\*", f"Aktueller Stand: **{VERSION}** · Letzte Aktualisierung: **{DATE_TEXT}**"),
        (r"RELEASE_NOTES_v\d+\.\d+\.\d+\.md", f"RELEASE_NOTES_{VERSION}.md"),
    ],
    ROOT / 'docs' / 'DOKUMENTATION_TECHNIK.md': [
        (r"Aktueller veröffentlichter Stand im Repository: `v\d+\.\d+\.\d+`", f"Aktueller veröffentlichter Stand im Repository: `{VERSION}`"),
        (r"RELEASE_NOTES_v\d+\.\d+\.\d+\.md", f"RELEASE_NOTES_{VERSION}.md"),
    ],
}

for file_path, replacements in REPLACEMENTS.items():
    if not file_path.exists():
        continue
    content = file_path.read_text(encoding='utf-8')
    updated = content
    for pattern, replacement in replacements:
        updated = re.sub(pattern, replacement, updated)
    if updated != content:
        file_path.write_text(updated, encoding='utf-8')
        print(f'aktualisiert: {file_path.relative_to(ROOT)}')
