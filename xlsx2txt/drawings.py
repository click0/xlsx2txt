"""Images and charts: extraction from the archive and (de)serialization."""

import hashlib
import posixpath
import zipfile
from io import BytesIO
from typing import Any

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
from openpyxl.packaging.relationship import get_dependents, get_rel, get_rels_path
from openpyxl.reader.workbook import WorkbookParser
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string
from openpyxl.xml.constants import IMAGE_NS
from openpyxl.xml.functions import fromstring, tostring

# ---------------------------------------------------------------------------
# Anchors
# ---------------------------------------------------------------------------


def _marker_to_json(marker: AnchorMarker) -> dict[str, Any]:
    result: dict[str, Any] = {"cell": f"{get_column_letter(marker.col + 1)}{marker.row + 1}"}
    if marker.colOff or marker.rowOff:
        result["offset"] = [marker.colOff, marker.rowOff]
    return result


def _marker_from_json(data: dict[str, Any]) -> AnchorMarker:
    letter, row = coordinate_from_string(data["cell"])
    col_off, row_off = data.get("offset", [0, 0])
    return AnchorMarker(col=column_index_from_string(letter) - 1, row=row - 1,
                        colOff=col_off, rowOff=row_off)


def anchor_to_json(anchor: Any) -> dict[str, Any]:
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


def anchor_from_json(data: dict[str, Any]) -> Any:
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


def extract_images(path) -> tuple[dict[str, list[tuple[bytes, str, Any]]], list[str]]:
    """Read images straight from the archive.

    Returns ``({sheet name: [(content, extension, anchor), ...]}, warnings)``.
    openpyxl only keeps images when Pillow is installed and loses the
    original bytes, so the archive is parsed directly.
    """
    result: dict[str, list[tuple[bytes, str, Any]]] = {}
    warnings: list[str] = []
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
                    warnings.append(
                        f"Sheet '{sheet.name}': {drawing_path} could not be read, "
                        "its images and charts are not exported"
                    )
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


def build_image(content: bytes, anchor: dict[str, Any]):
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

def chart_to_json(chart, source: dict[str, Any] | None = None) -> dict[str, Any]:
    """``source``: the chart as stored in the file (see extract_charts());
    without it the chart is written as openpyxl understands it."""
    if source and "xml" in source:
        result = {"anchor": anchor_to_json(chart.anchor), "xml": source["xml"]}
        if source["rels"]:
            result["rels"] = source["rels"]
        return result
    if source:
        chart.style = source["style"]
        chart.roundedCorners = source["roundedCorners"]
    return {
        "anchor": anchor_to_json(chart.anchor),
        "xml": tostring(chart._write()).decode("utf-8"),
    }


def chart_from_json(data: dict[str, Any]):
    space = ChartSpace.from_tree(fromstring(data["xml"]))
    chart = read_chart(space)
    # read_chart() drops these chart space settings.
    chart.style = space.style
    chart.roundedCorners = space.roundedCorners
    chart.anchor = anchor_from_json(data["anchor"])
    return chart


# Parts a chart may use that are kept with it: short name -> relationship type.
CHART_RELS = {
    "chartStyle": "http://schemas.microsoft.com/office/2011/relationships/chartStyle",
    "chartColorStyle": "http://schemas.microsoft.com/office/2011/relationships/chartColorStyle",
    "chartUserShapes": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartUserShapes",
    "package": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package",
    "image": IMAGE_NS,
}
# File name prefixes in data/charts/ (images go to data/media/).
_CHART_PART_PREFIX = {"chartStyle": "style", "chartColorStyle": "colors", "chartUserShapes": "shapes",
                      "package": "package"}


def chart_part_name(kind: str, content: bytes, extension: str) -> str:
    """Content-addressed file name of a part a chart uses."""
    return f"{_CHART_PART_PREFIX[kind]}_{hashlib.sha256(content).hexdigest()[:16]}.{extension.lower()}"


def _sheet_charts(archive: zipfile.ZipFile):
    """Yield ``(sheet name, chart part, ChartSpace)`` for every chart, in the
    order openpyxl reads them."""
    names = set(archive.namelist())
    parser = WorkbookParser(archive, "xl/workbook.xml")
    parser.parse()
    for sheet, rel in parser.find_sheets():
        rels_path = get_rels_path(rel.target)
        if rel.target not in names or rels_path not in names:
            continue
        for drawing_rel in get_dependents(archive, rels_path).find(SpreadsheetDrawing._rel_type):
            if drawing_rel.target not in names:
                continue
            try:
                drawing = SpreadsheetDrawing.from_tree(fromstring(archive.read(drawing_rel.target)))
            except TypeError:
                continue
            deps_path = get_rels_path(drawing_rel.target)
            deps = get_dependents(archive, deps_path) if deps_path in names else []
            for chart_rel in drawing._chart_rels:
                try:
                    space = get_rel(archive, deps, chart_rel.id, ChartSpace)
                except TypeError:
                    continue  # openpyxl skips such charts too
                yield sheet.name, deps.get(chart_rel.id).target, space


def _chart_source(archive: zipfile.ZipFile, names: set[str], part: str,
                  parts: dict[str, bytes], media: dict[str, bytes]) -> tuple[dict | None, str | None]:
    """The chart XML as stored and the parts it uses: ``({"xml", "rels"},
    None)``, or ``(None, reason)`` when it cannot be kept as it is."""
    try:
        xml = archive.read(part).decode("utf-8")
    except UnicodeDecodeError:
        return None, "its XML is not UTF-8"
    rels: dict[str, dict] = {}
    rels_path = get_rels_path(part)
    kinds = {value: key for key, value in CHART_RELS.items()}
    for rel in get_dependents(archive, rels_path) if rels_path in names else []:
        kind = kinds.get(rel.Type)
        if kind is None or rel.TargetMode == "External" or rel.target not in names:
            return None, f"it refers to {rel.Type.rsplit('/', 1)[-1]} {rel.Target}"
        if kind == "chartUserShapes" and get_rels_path(rel.target) in names \
                and len(get_dependents(archive, get_rels_path(rel.target))):
            return None, "its shapes refer to other parts"
        content = archive.read(rel.target)
        extension = posixpath.splitext(rel.target)[1].lstrip(".") or "xml"
        if kind == "image":
            name = media_name(content, extension)
            media[name] = content
        else:
            name = chart_part_name(kind, content, extension)
            parts[name] = content
        rels[rel.Id] = {"type": kind, "file": name}
    return {"xml": xml, "rels": rels}, None


def extract_charts(path) -> tuple[dict[str, list[dict[str, Any]]], dict[str, bytes], dict[str, bytes],
                                  set[str], list[str]]:
    """Read every chart as stored in the file.

    Returns ``({sheet name: [chart info, ...]}, chart parts, media, parts
    kept, warnings)`` with one entry per chart, in the order openpyxl reads
    them: ``{"xml", "rels"}`` (the chart XML byte for byte and the style,
    colors, shapes, embedded data and pictures it uses), or ``{"style",
    "roundedCorners"}`` for a chart that is exported as openpyxl reads it
    (its reader drops both settings; without roundedCorners="0" Excel draws
    rounded corners).
    """
    result: dict[str, list[dict[str, Any]]] = {}
    parts: dict[str, bytes] = {}
    media: dict[str, bytes] = {}
    kept: set[str] = set()
    warnings: list[str] = []
    with zipfile.ZipFile(path) as archive:
        names = set(archive.namelist())
        for sheet_name, part, space in _sheet_charts(archive):
            charts = result.setdefault(sheet_name, [])
            source, reason = _chart_source(archive, names, part, parts, media)
            if source is None:
                warnings.append(f"Sheet '{sheet_name}': chart {len(charts) + 1} is exported as openpyxl reads it "
                                f"({reason}); features openpyxl does not know are lost")
                charts.append({"style": space.style, "roundedCorners": space.roundedCorners})
                continue
            rels_path = get_rels_path(part)
            if rels_path in names:
                kept.update(rel.target for rel in get_dependents(archive, rels_path))
            charts.append(source)
    return result, parts, media, kept, warnings


def chart_title(data: dict[str, Any]) -> str | None:
    """Best-effort chart title for summaries."""
    try:
        tree = fromstring(data["xml"])
    except Exception:
        return None
    texts = [el.text for el in tree.iter() if el.tag.endswith("}t") and el.text]
    return texts[0] if texts else None
