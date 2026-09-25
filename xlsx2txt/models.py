"""Data models for xlsx2txt."""

from dataclasses import dataclass, field
from typing import Any


@dataclass
class FontStyle:
    """Font style properties."""
    name: str = "Calibri"
    size: float = 11.0
    bold: bool = False
    italic: bool = False
    underline: str | None = None
    strike: bool = False
    color: str | None = None


@dataclass
class FillStyle:
    """Cell fill/background properties."""
    pattern_type: str | None = None
    fg_color: str | None = None
    bg_color: str | None = None


@dataclass
class BorderSide:
    """Single border side properties."""
    style: str | None = None
    color: str | None = None


@dataclass
class BorderStyle:
    """Cell border properties."""
    left: BorderSide | None = None
    right: BorderSide | None = None
    top: BorderSide | None = None
    bottom: BorderSide | None = None


@dataclass
class AlignmentStyle:
    """Cell alignment properties."""
    horizontal: str | None = None
    vertical: str | None = None
    wrap_text: bool = False
    text_rotation: int = 0


@dataclass
class CellStyle:
    """Complete cell style."""
    font: FontStyle = field(default_factory=FontStyle)
    fill: FillStyle = field(default_factory=FillStyle)
    border: BorderStyle = field(default_factory=BorderStyle)
    alignment: AlignmentStyle = field(default_factory=AlignmentStyle)
    number_format: str = "General"


@dataclass
class Cell:
    """Cell with value and style."""
    coordinate: str
    value: Any = None
    data_type: str = "n"
    formula: str | None = None
    style: CellStyle = field(default_factory=CellStyle)


@dataclass
class ColumnDimension:
    """Column dimension properties."""
    width: float = 8.43
    hidden: bool = False
    style: int | None = None


@dataclass
class RowDimension:
    """Row dimension properties."""
    height: float = 15.0
    hidden: bool = False
    style: int | None = None


@dataclass
class SheetDimensions:
    """Sheet dimensions."""
    used_range: str | None = None
    columns: dict[str, ColumnDimension] = field(default_factory=dict)
    rows: dict[str, RowDimension] = field(default_factory=dict)
    default_row_height: float = 15.0
    default_col_width: float = 8.43


@dataclass
class CellData:
    """Compact cell data for JSON export."""
    v: Any = None  # value
    t: str = "n"   # type: s=string, n=number, b=boolean, d=date, e=error
    s: int | None = None  # style index
    f: str | None = None  # formula
    f_type: str | None = None  # formula type: array, shared
    f_ref: str | None = None  # formula reference range


@dataclass
class Sheet:
    """Complete sheet data."""
    sheet_id: int
    name: str
    dimensions: SheetDimensions = field(default_factory=SheetDimensions)
    merged_cells: list[str] = field(default_factory=list)
    cells: dict[str, CellData] = field(default_factory=dict)
    
    def to_dict(self) -> dict:
        """Convert to JSON-serializable dictionary."""
        return {
            "sheetId": self.sheet_id,
            "name": self.name,
            "dimensions": {
                "usedRange": self.dimensions.used_range,
                "columns": {
                    k: {"width": v.width, "hidden": v.hidden, "style": v.style}
                    for k, v in self.dimensions.columns.items()
                },
                "rows": {
                    k: {"height": v.height, "hidden": v.hidden, "style": v.style}
                    for k, v in self.dimensions.rows.items()
                },
                "defaultRowHeight": self.dimensions.default_row_height,
                "defaultColWidth": self.dimensions.default_col_width,
            },
            "mergedCells": self.merged_cells,
            "cells": {
                k: self._cell_to_dict(v)
                for k, v in self.cells.items()
            },
        }
    
    def _cell_to_dict(self, cell: CellData) -> dict:
        """Convert cell to compact dict."""
        result = {"v": cell.v, "t": cell.t}
        if cell.s is not None:
            result["s"] = cell.s
        if cell.f is not None:
            result["f"] = cell.f
        if cell.f_type is not None:
            result["fType"] = cell.f_type
        if cell.f_ref is not None:
            result["fRef"] = cell.f_ref
        return result
