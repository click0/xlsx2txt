"""The workbooks in examples/ must round-trip and match their committed exports."""

import importlib.util
from pathlib import Path

import pytest

from xlsx2txt import diff_models, export_model
from xlsx2txt.converter import roundtrip_diff
from xlsx2txt.storage import read_model, write_model
from xlsx2txt.text import render_text

EXAMPLES_DIR = Path(__file__).resolve().parent.parent / "examples"
EXAMPLE_FILES = sorted(EXAMPLES_DIR.glob("*.xlsx"))
REGENERATE = "run `python examples/generate.py` and commit the result"


def _load_generator():
    spec = importlib.util.spec_from_file_location("examples_generate", EXAMPLES_DIR / "generate.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def test_examples_exist():
    generator = _load_generator()
    assert [p.stem for p in EXAMPLE_FILES] == sorted(generator.EXAMPLES), REGENERATE


@pytest.mark.parametrize("path", EXAMPLE_FILES, ids=lambda p: p.stem)
def test_example_roundtrip(path):
    assert roundtrip_diff(export_model(path)) == []


@pytest.mark.parametrize("path", EXAMPLE_FILES, ids=lambda p: p.stem)
def test_committed_export_is_up_to_date(path, tmp_path):
    committed_dir = EXAMPLES_DIR / path.stem
    assert committed_dir.is_dir(), REGENERATE
    fresh = export_model(path)
    assert diff_models(read_model(committed_dir), fresh) == [], REGENERATE

    # Same files with the same content (manifest and checksums carry the version).
    write_model(fresh, tmp_path / "fresh")
    data_files = sorted(p.relative_to(committed_dir) for p in (committed_dir / "data").rglob("*") if p.is_file())
    fresh_files = sorted(p.relative_to(tmp_path / "fresh") for p in (tmp_path / "fresh" / "data").rglob("*") if p.is_file())
    assert data_files == fresh_files, REGENERATE
    for rel in data_files:
        assert (committed_dir / rel).read_bytes() == (tmp_path / "fresh" / rel).read_bytes(), f"{rel}: {REGENERATE}"


def test_generator_matches_committed_workbooks(tmp_path):
    generator = _load_generator()
    for generated in generator.generate(tmp_path, export=False):
        committed = EXAMPLES_DIR / generated.name
        assert diff_models(export_model(committed), export_model(generated)) == [], f"{generated.name}: {REGENERATE}"


@pytest.mark.parametrize("path", EXAMPLE_FILES, ids=lambda p: p.stem)
def test_example_renders_as_text(path):
    text = render_text(export_model(path))
    assert text.startswith(("=== ", "name ", "external link"))
