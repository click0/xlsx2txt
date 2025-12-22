"""Data models for xlsx2txt."""

from dataclasses import dataclass, field
from typing import Any, Optional


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
