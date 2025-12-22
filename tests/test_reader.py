"""Tests for xlsx2txt.reader module."""

import json
import pytest
from pathlib import Path

from xlsx2txt import read_cell, read_cell_style, read_sheet
from xlsx2txt.models import Cell, CellStyle, Sheet


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


class TestReadSheet:
    """Tests for read_sheet function."""

    def test_read_simple_sheet(self):
        """Read simple sheet."""
        sheet = read_sheet(FIXTURES_DIR / "simple.xlsx")

        assert isinstance(sheet, Sheet)
        assert sheet.name == "Sheet"
        assert sheet.sheet_id == 1
        assert "A1" in sheet.cells
        assert sheet.cells["A1"].v == "Hello"
        assert sheet.cells["A1"].t == "s"

    def test_read_sheet_with_formulas(self):
        """Read sheet with formulas."""
        sheet = read_sheet(FIXTURES_DIR / "formula.xlsx")

        assert "A1" in sheet.cells
        assert sheet.cells["A1"].v == 10
        assert sheet.cells["A2"].v == 20
        assert sheet.cells["A3"].f == "=A1+A2"

    def test_read_complex_sheet(self):
        """Read complex sheet with merged cells and various data."""
        sheet = read_sheet(FIXTURES_DIR / "complex.xlsx", "Data")

        # Check merged cells
        assert "A1:D1" in sheet.merged_cells

        # Check header
        assert sheet.cells["A1"].v == "Sales Report Q1 2025"

        # Check data
        assert sheet.cells["A3"].v == "Widget"
        assert sheet.cells["B3"].v == 100
        assert sheet.cells["C3"].v == 9.99
        assert sheet.cells["D3"].f == "=B3*C3"

        # Check dimensions
        assert "A" in sheet.dimensions.columns
        assert sheet.dimensions.columns["A"].width == 15

    def test_sheet_to_dict(self):
        """Test JSON serialization."""
        sheet = read_sheet(FIXTURES_DIR / "simple.xlsx")
        data = sheet.to_dict()

        assert data["name"] == "Sheet"
        assert data["sheetId"] == 1
        assert "cells" in data
        assert "A1" in data["cells"]
        assert data["cells"]["A1"]["v"] == "Hello"

        # Should be JSON serializable
        json_str = json.dumps(data)
        assert len(json_str) > 0

    def test_file_not_found(self):
        """Raise error for non-existent file."""
        with pytest.raises(FileNotFoundError):
            read_sheet(FIXTURES_DIR / "nonexistent.xlsx")

    def test_invalid_sheet_name(self):
        """Raise error for invalid sheet name."""
        with pytest.raises(ValueError, match="Sheet not found"):
            read_sheet(FIXTURES_DIR / "simple.xlsx", "NonExistent")
