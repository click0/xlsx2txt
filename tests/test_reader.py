"""Tests for xlsx2txt.reader module."""

import pytest
from pathlib import Path

from xlsx2txt import read_cell, read_cell_style
from xlsx2txt.models import Cell, CellStyle


FIXTURES_DIR = Path(__file__).parent / "fixtures"


class TestReadCell:
    """Tests for read_cell function."""

    def test_read_simple_cell(self):
        """Read plain text from A1."""
        cell = read_cell(FIXTURES_DIR / "simple.xlsx", "A1")

        assert isinstance(cell, Cell)
        assert cell.coordinate == "A1"
        assert cell.value == "Hello"
        assert cell.data_type == "s"
        assert cell.formula is None

    def test_read_styled_cell_value(self):
        """Read value from styled A1."""
        cell = read_cell(FIXTURES_DIR / "styled.xlsx", "A1")

        assert cell.value == "Styled"
        assert cell.data_type == "s"

    def test_read_empty_cell(self):
        """Read empty A1."""
        cell = read_cell(FIXTURES_DIR / "empty.xlsx", "A1")

        assert cell.value is None
        assert cell.data_type == "n"

    def test_read_formula_cell(self):
        """Read cell with formula."""
        cell = read_cell(FIXTURES_DIR / "formula.xlsx", "A3")

        assert cell.formula == "=A1+A2"

    def test_file_not_found(self):
        """Raise error for non-existent file."""
        with pytest.raises(FileNotFoundError):
            read_cell(FIXTURES_DIR / "nonexistent.xlsx", "A1")

    def test_invalid_sheet(self):
        """Raise error for invalid sheet name."""
        with pytest.raises(ValueError, match="Sheet not found"):
            read_cell(FIXTURES_DIR / "simple.xlsx", "A1", sheet_name="NonExistent")


class TestReadCellStyle:
    """Tests for read_cell_style function."""

    def test_read_simple_style(self):
        """Read default style from simple cell."""
        style = read_cell_style(FIXTURES_DIR / "simple.xlsx", "A1")

        assert isinstance(style, CellStyle)
        assert style.font.name == "Calibri"
        assert style.font.bold is False

    def test_read_styled_font(self):
        """Read font style from styled cell."""
        style = read_cell_style(FIXTURES_DIR / "styled.xlsx", "A1")

        assert style.font.name == "Arial"
        assert style.font.size == 14
        assert style.font.bold is True
        assert style.font.color == "00FF0000"  # Red

    def test_read_styled_fill(self):
        """Read fill style from styled cell."""
        style = read_cell_style(FIXTURES_DIR / "styled.xlsx", "A1")

        assert style.fill.pattern_type == "solid"
        assert style.fill.fg_color == "00FFFF00"  # Yellow

    def test_read_styled_border(self):
        """Read border style from styled cell."""
        style = read_cell_style(FIXTURES_DIR / "styled.xlsx", "A1")

        assert style.border.left is not None
        assert style.border.left.style == "thin"
        assert style.border.right is not None
        assert style.border.top is not None
        assert style.border.bottom is not None

    def test_read_styled_alignment(self):
        """Read alignment style from styled cell."""
        style = read_cell_style(FIXTURES_DIR / "styled.xlsx", "A1")

        assert style.alignment.horizontal == "center"
        assert style.alignment.vertical == "center"

    def test_empty_cell_style(self):
        """Read style from empty cell (should have defaults)."""
        style = read_cell_style(FIXTURES_DIR / "empty.xlsx", "A1")

        assert isinstance(style, CellStyle)
        assert style.font.name == "Calibri"
