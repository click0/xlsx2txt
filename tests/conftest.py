"""Shared pytest fixtures."""

from pathlib import Path

import pytest

from tests.create_fixtures import main as create_fixtures
from tests.rich_workbook import build_rich_workbook

FIXTURES_DIR = Path(__file__).parent / "fixtures"
EXAMPLES_DIR = Path(__file__).parent.parent / "examples"
# A richer workbook used by the CLI tests; it lives with the examples.
COMPLEX_XLSX = EXAMPLES_DIR / "10-complex.xlsx"
REQUIRED = ["simple.xlsx", "styled.xlsx", "empty.xlsx", "formula.xlsx"]


def pytest_sessionstart(session):
    """Generate fixture files if some of them are missing."""
    if not all((FIXTURES_DIR / name).exists() for name in REQUIRED):
        create_fixtures()


@pytest.fixture
def rich_xlsx(tmp_path):
    return build_rich_workbook(tmp_path / "rich.xlsx")
