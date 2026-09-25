"""Build the GitHub release text for a version from the changelogs.

Usage: python scripts/release_notes.py 0.3.0 > notes.md

Exits with an error when a changelog has no section for the version, so a
release cannot be published without a description.
"""

import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
REPO = "https://github.com/click0/xlsx2txt"


def section(changelog: Path, version: str) -> str:
    """Return the body of '## [version]' up to the next '## ' heading."""
    text = changelog.read_text(encoding="utf-8")
    match = re.search(rf"^## \[{re.escape(version)}\][^\n]*\n(.*?)(?=^## |^\[[^\]]+\]: |\Z)",
                      text, re.MULTILINE | re.DOTALL)
    if not match or not match.group(1).strip():
        raise SystemExit(f"{changelog.name}: no section for version {version}")
    return match.group(1).strip()


def notes(version: str) -> str:
    english = section(ROOT / "CHANGELOG.md", version)
    ukrainian = section(ROOT / "CHANGELOG_UK.md", version)
    wheel = f"{REPO}/releases/download/v{version}/xlsx2txt-{version}-py3-none-any.whl"
    return f"""{english}

### Installation

```bash
pip install {wheel}
# with VBA source extraction
pip install "xlsx2txt[vba] @ {wheel}"
```

<details>
<summary>Українською</summary>

{ukrainian}

### Встановлення

```bash
pip install {wheel}
```

</details>
"""


if __name__ == "__main__":
    if len(sys.argv) != 2:
        raise SystemExit("usage: release_notes.py VERSION")
    sys.stdout.write(notes(sys.argv[1].lstrip("v")))
