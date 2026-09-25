"""Images and charts: extraction from the archive and (de)serialization."""

import hashlib
import posixpath
import zipfile
from io import BytesIO
from typing import Any, Dict, List, Optional, Tuple

from openpyxl.chart.chartspace import ChartSpace
from openpyxl.chart.reader import read_chart
from openpyxl.drawing.spreadsheet_drawing import (
    AbsoluteAnchor,
    AnchorMarker,
    OneCellAnchor,
    SpreadsheetDrawing,
    TwoCellAnchor,
)
from openpyxl.drawing.xdr import XDRPoint2D, XDRPositiveSize2D
from openpyxl.packaging.relationship import get_dependents, get_rels_path
from openpyxl.reader.workbook import WorkbookParser
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string
from openpyxl.xml.constants import IMAGE_NS
from openpyxl.xml.functions import fromstring, tostring

# ---------------------------------------------------------------------------
# Anchors
# ---------------------------------------------------------------------------


def _marker_to_json(marker: AnchorMarker) -> Dict[str, Any]:
    result: Dict[str, Any] = {"cell": f"{get_column_letter(marker.col + 1)}{marker.row + 1}"}
    if marker.colOff or marker.rowOff:
        result["offset"] = [marker.colOff, marker.rowOff]
    return result


def _marker_from_json(data: Dict[str, Any]) -> AnchorMarker:
    letter, row = coordinate_from_string(data["cell"])
    col_off, row_off = data.get("offset", [0, 0])
    return AnchorMarker(col=column_index_from_string(letter) - 1, row=row - 1,
                        colOff=col_off, rowOff=row_off)


def anchor_to_json(anchor: Any) -> Dict[str, Any]:
    """Describe where a drawing object sits; sizes are in EMU."""
    if isinstance(anchor, str):
        return {"type": "oneCell", "from": {"cell": anchor}}
    if isinstance(anchor, TwoCellAnchor):
        result = {"type": "twoCell", "from": _marker_to_json(anchor._from), "to": _marker_to_json(anchor.to)}
        if anchor.editAs:
            result["editAs"] = anchor.editAs
        return result
    if isinstance(anchor, OneCellAnchor):
        result = {"type": "oneCell", "from": _marker_to_json(anchor._from)}
        if anchor.ext is not None:
            result["size"] = [anchor.ext.width, anchor.ext.height]
        return result
    if isinstance(anchor, AbsoluteAnchor):
        return {
            "type": "absolute",
            "position": [anchor.pos.x, anchor.pos.y],
            "size": [anchor.ext.width, anchor.ext.height],
        }
    raise ValueError(f"Unsupported anchor: {type(anchor).__name__}")


def anchor_from_json(data: Dict[str, Any]) -> Any:
    kind = data.get("type")
    if kind == "twoCell":
        return TwoCellAnchor(_from=_marker_from_json(data["from"]), to=_marker_from_json(data["to"]),
                             editAs=data.get("editAs"))
    if kind == "oneCell":
        ext = None
        if "size" in data:
            ext = XDRPositiveSize2D(cx=data["size"][0], cy=data["size"][1])
        return OneCellAnchor(_from=_marker_from_json(data["from"]), ext=ext)
    if kind == "absolute":
        return AbsoluteAnchor(pos=XDRPoint2D(x=data["position"][0], y=data["position"][1]),
                              ext=XDRPositiveSize2D(cx=data["size"][0], cy=data["size"][1]))
    raise ValueError(f"Unsupported anchor type: {kind!r}")


# ---------------------------------------------------------------------------
# Images
# ---------------------------------------------------------------------------

def media_name(content: bytes, extension: str) -> str:
    """Content-addressed file name, stable across exports."""
    return f"image_{hashlib.sha256(content).hexdigest()[:16]}.{extension.lower()}"


def extract_images(path) -> Tuple[Dict[str, List[Tuple[bytes, str, Any]]], List[str]]:
    """Read images straight from the archive.

    Returns ``({sheet name: [(content, extension, anchor), ...]}, warnings)``.
    openpyxl only keeps images when Pillow is installed and loses the
    original bytes, so the archive is parsed directly.
    """
    result: Dict[str, List[Tuple[bytes, str, Any]]] = {}
    warnings: List[str] = []
    with zipfile.ZipFile(path) as archive:
        names = set(archive.namelist())
        parser = WorkbookParser(archive, "xl/workbook.xml")
        parser.parse()
        for sheet, rel in parser.find_sheets():
            sheet_path = rel.target
            rels_path = get_rels_path(sheet_path)
            if sheet_path not in names or rels_path not in names:
                continue
            for drawing_rel in get_dependents(archive, rels_path).find(SpreadsheetDrawing._rel_type):
                drawing_path = drawing_rel.target
                if drawing_path not in names:
                    continue
                try:
                    drawing = SpreadsheetDrawing.from_tree(fromstring(archive.read(drawing_path)))
                except TypeError:
                    warnings.append(f"Sheet '{sheet.name}': shapes in {drawing_path} are not exported")
                    continue
                drawing_rels_path = get_rels_path(drawing_path)
                if drawing_rels_path not in names:
                    continue
                deps = get_dependents(archive, drawing_rels_path)
                for blip in drawing._blip_rels:
                    dep = deps.get(blip.embed)
                    if dep is None or dep.Type != IMAGE_NS or dep.target not in names:
                        continue
                    extension = posixpath.splitext(dep.target)[1].lstrip(".") or "png"
                    if extension.lower() in ("wmf", "emf"):
                        warnings.append(f"Sheet '{sheet.name}': {extension.upper()} image {dep.target} is not supported")
                        continue
                    result.setdefault(sheet.name, []).append(
                        (archive.read(dep.target), extension, blip.anchor)
                    )
    return result, warnings


def build_image(content: bytes, anchor: Dict[str, Any]):
    """Create an openpyxl image (requires Pillow)."""
    try:
        from openpyxl.drawing.image import Image
        image = Image(BytesIO(content))
    except ImportError:
        raise RuntimeError("Pillow is required to import images: pip install pillow")
    image.anchor = anchor_from_json(anchor)
    # Keep the displayed size of one-cell anchors.
    if anchor.get("type") == "oneCell" and "size" in anchor:
        image.width = anchor["size"][0] / 9525
        image.height = anchor["size"][1] / 9525
    return image


# ---------------------------------------------------------------------------
# Charts
# ---------------------------------------------------------------------------

def chart_to_json(chart) -> Dict[str, Any]:
    return {
        "anchor": anchor_to_json(chart.anchor),
        "xml": tostring(chart._write()).decode("utf-8"),
    }


def chart_from_json(data: Dict[str, Any]):
    chart = read_chart(ChartSpace.from_tree(fromstring(data["xml"])))
    chart.anchor = anchor_from_json(data["anchor"])
    return chart


def chart_title(data: Dict[str, Any]) -> Optional[str]:
    """Best-effort chart title for summaries."""
    try:
        tree = fromstring(data["xml"])
    except Exception:
        return None
    texts = [el.text for el in tree.iter() if el.tag.endswith("}t") and el.text]
    return texts[0] if texts else None
