from __future__ import annotations

from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
TARGETS = [
    ROOT / 'README.md',
    ROOT / 'docs' / 'CHANGELOG.md',
]

TARGETS.extend(sorted((ROOT / 'release').glob('RELEASE_NOTES_v*.md')))

for doc in TARGETS:
    if not doc.exists():
        continue
    text = doc.read_text(encoding='utf-8').replace('\r\n', '\n').rstrip() + '\n'
    doc.write_text(text, encoding='utf-8')
    print(f'normalisiert: {doc.relative_to(ROOT)}')
