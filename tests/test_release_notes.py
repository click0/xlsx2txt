"""Release descriptions come from CHANGELOG.md / CHANGELOG_UK.md."""

import importlib.util
from pathlib import Path

import pytest

import xlsx2txt

ROOT = Path(__file__).resolve().parent.parent


def _script():
    spec = importlib.util.spec_from_file_location("release_notes", ROOT / "scripts" / "release_notes.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def test_current_version_has_release_notes():
    """Bumping __version__ requires a changelog section in both languages."""
    text = _script().notes(xlsx2txt.__version__)
    assert f"v{xlsx2txt.__version__}/xlsx2txt-{xlsx2txt.__version__}-py3-none-any.whl" in text
    assert "<summary>Українською</summary>" in text
    # Only this version's section: no other "## [x.y.z]" headings.
    assert not [line for line in text.splitlines() if line.startswith("## ")]


@pytest.mark.parametrize("version", ["0.1.0", "0.2.0", "0.3.0"])
def test_sections_are_separate(version):
    text = _script().notes(version)
    assert "[Unreleased]" not in text
    assert "compare/" not in text  # link references at the bottom are not included


def test_missing_version_fails():
    with pytest.raises(SystemExit, match="no section for version 9.9.9"):
        _script().notes("9.9.9")
