"""Conversion of openpyxl style objects to/from JSON-friendly dicts."""

import json
from typing import Any

from openpyxl.styles import Alignment, Border, Font, PatternFill, GradientFill, Protection, Side
from openpyxl.styles.colors import Color
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.styles.fills import Stop

# ---------------------------------------------------------------------------
# Colors
# ---------------------------------------------------------------------------

_DEFAULT_RGB = "00000000"


def color_to_json(color: Color | None) -> Any:
    """Serialize a color.

    Plain RGB colors become a string ("FFFF0000"); theme, indexed and
    automatic colors (or any color with a tint) become a small dict.
    """
    if color is None:
        return None
    tint = color.tint or 0
    if color.type == "rgb":
        if not isinstance(color.rgb, str):
            return None
        if not tint:
            return color.rgb
        result = {"rgb": color.rgb}
    elif color.type == "theme":
        result = {"theme": color.theme}
    elif color.type == "indexed":
        result = {"indexed": color.indexed}
    elif color.type == "auto":
        result = {"auto": True}
    else:
        return None
    if tint:
        result["tint"] = tint
    return result


def color_from_json(data: Any) -> Color | None:
    """Deserialize a color produced by :func:`color_to_json`."""
    if data is None:
        return None
    if isinstance(data, str):
        return Color(rgb=data)
    return Color(**data)


def _is_default_color(color: Color | None) -> bool:
    return (
        color is None
        or (color.type == "rgb" and color.rgb == _DEFAULT_RGB and not color.tint)
    )


# ---------------------------------------------------------------------------
# Individual style parts
# ---------------------------------------------------------------------------

_FONT_ATTRS = [
    "name", "size", "bold", "italic", "underline", "strike", "vertAlign",
    "color", "family", "charset", "scheme", "outline", "shadow", "condense", "extend",
]


def font_to_json(font: Font) -> dict[str, Any]:
    result = {}
    for attr in _FONT_ATTRS:
        value = getattr(font, attr)
        if attr == "color":
            value = color_to_json(value)
        if value is None or value is False:
            continue
        result[attr] = value
    return result


def font_from_json(data: dict[str, Any]) -> Font:
    kwargs = dict(data)
    if "color" in kwargs:
        kwargs["color"] = color_from_json(kwargs["color"])
    return Font(**kwargs)


def _unwrap(obj):
    """Return the object behind an openpyxl StyleProxy."""
    return getattr(obj, "_StyleProxy__target", obj)


def fill_to_json(fill) -> dict[str, Any]:
    fill = _unwrap(fill)
    if isinstance(fill, GradientFill):
        result = {"gradient": fill.type or "linear"}
        for attr in ("degree", "left", "right", "top", "bottom"):
            value = getattr(fill, attr)
            if value:
                result[attr] = value
        result["stops"] = [
            {"position": stop.position, "color": color_to_json(stop.color)}
            for stop in fill.stop
        ]
        return result
    result = {"patternType": fill.patternType}
    if not _is_default_color(fill.fgColor):
        result["fgColor"] = color_to_json(fill.fgColor)
    if not _is_default_color(fill.bgColor):
        result["bgColor"] = color_to_json(fill.bgColor)
    return result


def fill_from_json(data: dict[str, Any]):
    if "gradient" in data:
        stops = [
            Stop(color=color_from_json(s["color"]), position=s["position"])
            for s in data.get("stops", [])
        ]
        return GradientFill(
            type=data["gradient"],
            degree=data.get("degree", 0),
            left=data.get("left", 0),
            right=data.get("right", 0),
            top=data.get("top", 0),
            bottom=data.get("bottom", 0),
            stop=stops,
        )
    kwargs = {"patternType": data.get("patternType")}
    if "fgColor" in data:
        kwargs["fgColor"] = color_from_json(data["fgColor"])
    if "bgColor" in data:
        kwargs["bgColor"] = color_from_json(data["bgColor"])
    return PatternFill(**kwargs)


_BORDER_SIDES = ["left", "right", "top", "bottom", "diagonal", "vertical", "horizontal"]


def border_to_json(border: Border) -> dict[str, Any]:
    result = {}
    for name in _BORDER_SIDES:
        side = getattr(border, name)
        if side is None or (side.style is None and side.color is None):
            continue
        side_data = {"style": side.style}
        if side.color is not None:
            side_data["color"] = color_to_json(side.color)
        result[name] = side_data
    if border.diagonalUp:
        result["diagonalUp"] = True
    if border.diagonalDown:
        result["diagonalDown"] = True
    if border.outline is False:
        result["outline"] = False
    return result


def border_from_json(data: dict[str, Any]) -> Border:
    kwargs = {}
    for name in _BORDER_SIDES:
        if name in data:
            side = data[name]
            kwargs[name] = Side(style=side.get("style"), color=color_from_json(side.get("color")))
    for flag in ("diagonalUp", "diagonalDown", "outline"):
        if flag in data:
            kwargs[flag] = data[flag]
    return Border(**kwargs)


_ALIGNMENT_ATTRS = [
    "horizontal", "vertical", "textRotation", "wrapText", "shrinkToFit",
    "indent", "relativeIndent", "justifyLastLine", "readingOrder",
]


def alignment_to_json(alignment: Alignment) -> dict[str, Any]:
    result = {}
    for attr in _ALIGNMENT_ATTRS:
        value = getattr(alignment, attr)
        if value is None or value is False or value == 0:
            continue
        result[attr] = value
    return result


def alignment_from_json(data: dict[str, Any]) -> Alignment:
    return Alignment(**data)


def protection_to_json(protection: Protection) -> dict[str, Any]:
    return {"locked": bool(protection.locked), "hidden": bool(protection.hidden)}


def protection_from_json(data: dict[str, Any]) -> Protection:
    return Protection(locked=data.get("locked", True), hidden=data.get("hidden", False))


# ---------------------------------------------------------------------------
# Differential styles (used by conditional formatting)
# ---------------------------------------------------------------------------

def dxf_to_json(dxf: DifferentialStyle | None) -> dict[str, Any] | None:
    if dxf is None:
        return None
    result = {}
    if dxf.font is not None:
        result["font"] = font_to_json(dxf.font)
    if dxf.fill is not None:
        result["fill"] = fill_to_json(dxf.fill)
    if dxf.border is not None:
        result["border"] = border_to_json(dxf.border)
    if dxf.alignment is not None:
        result["alignment"] = alignment_to_json(dxf.alignment)
    if dxf.numFmt is not None:
        result["numFmt"] = {"id": dxf.numFmt.numFmtId, "code": dxf.numFmt.formatCode}
    if dxf.protection is not None:
        result["protection"] = protection_to_json(dxf.protection)
    return result


def dxf_from_json(data: dict[str, Any] | None) -> DifferentialStyle | None:
    if data is None:
        return None
    from openpyxl.styles.numbers import NumberFormat

    kwargs = {}
    if "font" in data:
        kwargs["font"] = font_from_json(data["font"])
    if "fill" in data:
        kwargs["fill"] = fill_from_json(data["fill"])
    if "border" in data:
        kwargs["border"] = border_from_json(data["border"])
    if "alignment" in data:
        kwargs["alignment"] = alignment_from_json(data["alignment"])
    if "numFmt" in data:
        kwargs["numFmt"] = NumberFormat(numFmtId=data["numFmt"]["id"], formatCode=data["numFmt"]["code"])
    if "protection" in data:
        kwargs["protection"] = protection_from_json(data["protection"])
    return DifferentialStyle(**kwargs)


# ---------------------------------------------------------------------------
# Style table
# ---------------------------------------------------------------------------

STYLE_PARTS = ["fonts", "fills", "borders", "alignments", "protections"]


class _Pool:
    """Ordered, de-duplicated list of JSON-serializable items."""

    def __init__(self):
        self.items: list[Any] = []
        self._index: dict[str, int] = {}

    def add(self, item: Any) -> int:
        key = json.dumps(item, sort_keys=True, ensure_ascii=False)
        if key not in self._index:
            self._index[key] = len(self.items)
            self.items.append(item)
        return self._index[key]


class StyleTable:
    """Collects cell styles while exporting.

    Every distinct style is stored once; cells reference it by index.
    Index 0 is always the workbook default style.
    """

    def __init__(self):
        self._pools = {name: _Pool() for name in STYLE_PARTS}
        self._styles = _Pool()

    def add(self, obj) -> int:
        """Register the style of a cell (or row/column dimension)."""
        style = {
            "font": self._pools["fonts"].add(font_to_json(obj.font)),
            "fill": self._pools["fills"].add(fill_to_json(obj.fill)),
            "border": self._pools["borders"].add(border_to_json(obj.border)),
            "alignment": self._pools["alignments"].add(alignment_to_json(obj.alignment)),
            "protection": self._pools["protections"].add(protection_to_json(obj.protection)),
            "numFmt": obj.number_format or "General",
        }
        return self._styles.add(style)

    def to_json(self) -> dict[str, list[Any]]:
        result = {name: pool.items for name, pool in self._pools.items()}
        result["cellStyles"] = self._styles.items
        return result


class StyleApplier:
    """Applies exported styles to cells while importing."""

    def __init__(self, styles: dict[str, list[Any]]):
        self._styles = styles
        self._objects: dict[int, tuple] = {}
        self._arrays: dict[int, Any] = {}

    def objects(self, index: int) -> tuple:
        if index not in self._objects:
            style = self._styles["cellStyles"][index]
            self._objects[index] = (
                font_from_json(self._styles["fonts"][style["font"]]),
                fill_from_json(self._styles["fills"][style["fill"]]),
                border_from_json(self._styles["borders"][style["border"]]),
                alignment_from_json(self._styles["alignments"][style["alignment"]]),
                protection_from_json(self._styles["protections"][style["protection"]]),
                style.get("numFmt", "General"),
            )
        return self._objects[index]

    def apply(self, obj, index: int) -> None:
        from copy import copy

        if index in self._arrays:
            obj._style = copy(self._arrays[index])
            return
        font, fill, border, alignment, protection, num_fmt = self.objects(index)
        obj.font = font
        obj.fill = fill
        obj.border = border
        obj.alignment = alignment
        obj.protection = protection
        obj.number_format = num_fmt
        self._arrays[index] = copy(obj._style)


# ---------------------------------------------------------------------------
# Rich text (several differently formatted runs inside one cell)
# ---------------------------------------------------------------------------

_INLINE_FONT_ATTRS = [
    "rFont", "sz", "b", "i", "u", "strike", "vertAlign", "color",
    "family", "charset", "scheme", "outline", "shadow", "condense", "extend",
]


def rich_text_to_json(value) -> list[Any]:
    """Plain runs become strings, formatted runs {"text": ..., "font": {...}}."""
    runs: list[Any] = []
    for run in value:
        if isinstance(run, str):
            runs.append(run)
            continue
        font = {}
        for attr in _INLINE_FONT_ATTRS:
            attr_value = getattr(run.font, attr, None)
            if attr == "color":
                attr_value = color_to_json(attr_value)
            if attr_value is None or attr_value is False:
                continue
            font[attr] = attr_value
        runs.append({"text": run.text, "font": font})
    return runs


def rich_text_from_json(runs: list[Any]):
    from openpyxl.cell.rich_text import CellRichText, TextBlock
    from openpyxl.cell.text import InlineFont

    parts = []
    for run in runs:
        if isinstance(run, str):
            parts.append(run)
            continue
        kwargs = dict(run.get("font", {}))
        if "color" in kwargs:
            kwargs["color"] = color_from_json(kwargs["color"])
        parts.append(TextBlock(InlineFont(**kwargs), run["text"]))
    return CellRichText(parts)
