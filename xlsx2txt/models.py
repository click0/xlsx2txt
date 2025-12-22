"""Data models for xlsx2txt."""

from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional


@dataclass
class FontStyle:
    """Font style properties."""
    name: str = "Calibri"
    size: float = 11.0
    bold: bool = False
    italic: bool = False
    underline: Optional[str] = None
    strike: bool = False
    color: Optional[str] = None


@dataclass
class FillStyle:
    """Cell fill/background properties."""
    pattern_type: Optional[str] = None
    fg_color: Optional[str] = None
    bg_color: Optional[str] = None


@dataclass
class BorderSide:
    """Single border side properties."""
    style: Optional[str] = None
    color: Optional[str] = None


@dataclass
class BorderStyle:
    """Cell border properties."""
    left: Optional[BorderSide] = None
    right: Optional[BorderSide] = None
    top: Optional[BorderSide] = None
    bottom: Optional[BorderSide] = None


@dataclass
class AlignmentStyle:
    """Cell alignment properties."""
    horizontal: Optional[str] = None
    vertical: Optional[str] = None
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
    formula: Optional[str] = None
    style: CellStyle = field(default_factory=CellStyle)


@dataclass
class ColumnDimension:
    """Column dimension properties."""
    width: float = 8.43
    hidden: bool = False
    style: Optional[int] = None


@dataclass
class RowDimension:
    """Row dimension properties."""
    height: float = 15.0
    hidden: bool = False
    style: Optional[int] = None


@dataclass
class SheetDimensions:
    """Sheet dimensions."""
    used_range: Optional[str] = None
    columns: Dict[str, ColumnDimension] = field(default_factory=dict)
    rows: Dict[str, RowDimension] = field(default_factory=dict)
    default_row_height: float = 15.0
    default_col_width: float = 8.43


@dataclass
class CellData:
    """Compact cell data for JSON export."""
    v: Any = None  # value
    t: str = "n"   # type: s=string, n=number, b=boolean, d=date, e=error
    s: Optional[int] = None  # style index
    f: Optional[str] = None  # formula
    f_type: Optional[str] = None  # formula type: array, shared
    f_ref: Optional[str] = None  # formula reference range


@dataclass
class Sheet:
    """Complete sheet data."""
    sheet_id: int
    name: str
    dimensions: SheetDimensions = field(default_factory=SheetDimensions)
    merged_cells: List[str] = field(default_factory=list)
    cells: Dict[str, CellData] = field(default_factory=dict)
    
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
