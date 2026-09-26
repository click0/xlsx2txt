"""Shapes (text boxes, arrows, rectangles, connectors, groups) of worksheets.

openpyxl reads only pictures and charts from drawings, so shapes are taken
from the drawing XML directly: each drawing anchor holding a shape is kept
as a self-contained XML fragment and written back into the drawing of the
saved file on import.

Relationships a shape uses (a picture fill, a hyperlink) are stored next to
it under ``rels``; pictures go to ``data/media/``. ``layer`` is the number of
pictures and charts drawn below the shape (omitted when it is above all).
"""

import html
import posixpath
import re
from typing import Any, Callable

from xlsx2txt.drawings import media_name
from xlsx2txt.xmlfrag import self_contained as _self_contained
from xlsx2txt.xmlfrag import split_children as drawing_children

REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

_ANCHORS = {"twoCellAnchor", "oneCellAnchor", "absoluteAnchor", "AlternateContent"}
# The first graphic object in an anchor decides what it is.
_OBJECT = re.compile(r"<(?:\w+:)?(sp|grpSp|cxnSp|pic|graphicFrame|contentPart)\b")
_SHAPE_KINDS = {"sp", "grpSp", "cxnSp"}
# Form controls and legacy VML pictures depend on parts that are not exported.
_UNSUPPORTED = re.compile(r"\br:pict=|compatExt")
# References to relationships of the drawing (picture fills, hyperlinks).
_REL_ATTR = re.compile(r"\br:(embed|link|id)=\"([^\"]*)\"")
_ID_TAG = re.compile(r"<(?:\w+:)?(cNvPr|stCxn|endCxn)\b[^>]*>")
_ID_ATTR = re.compile(r"\bid=\"(\d+)\"")


def _local(name: str) -> str:
    return name.rsplit(":", 1)[-1]


def classify(fragment: str) -> str | None:
    """Return "shape", "unsupported shape" (form controls) or None."""
    match = re.match(r"<([\w.:-]+)", fragment)
    if not match or _local(match.group(1)) not in _ANCHORS:
        return None
    obj = _OBJECT.search(fragment)
    if not obj or obj.group(1) not in _SHAPE_KINDS:
        return None
    return "unsupported shape" if _UNSUPPORTED.search(fragment) else "shape"


def _is_graphic(fragment: str) -> bool:
    """A picture or chart anchor (what openpyxl writes back itself)."""
    obj = _OBJECT.search(fragment)
    return bool(obj) and (obj.group(1) == "pic"
                          or (obj.group(1) == "graphicFrame" and "drawingml/2006/chart" in fragment))


def _short_type(rel_type: str) -> str:
    return rel_type[len(REL_NS) + 1:] if rel_type.startswith(REL_NS + "/") else rel_type


def full_type(rel_type: str) -> str:
    return rel_type if "://" in rel_type else f"{REL_NS}/{rel_type}"


def _shape_rels(fragment: str, rels: dict, read: Callable[[str], bytes],
                media: dict[str, bytes]) -> dict[str, dict] | None:
    """Describe the relationships a shape uses; None if one cannot be kept."""
    result: dict[str, dict] = {}
    for _, rel_id in _REL_ATTR.findall(fragment):
        if not rel_id or rel_id in result:
            continue
        rel = rels.get(rel_id)
        if rel is None:
            return None
        rel_type = _short_type(rel.Type)
        if rel.TargetMode == "External":
            result[rel_id] = {"type": rel_type, "target": rel.Target, "external": True}
        elif rel_type == "image":
            try:
                content = read(rel.target)
            except KeyError:
                return None
            name = media_name(content, posixpath.splitext(rel.target)[1].lstrip(".") or "png")
            media[name] = content
            result[rel_id] = {"type": rel_type, "file": name}
        else:
            return None
    return result


def _text(fragment: str) -> str:
    paragraphs = re.split(r"</(?:\w+:)?p>", fragment)
    lines = ["".join(re.findall(r"<(?:\w+:)?t>([^<]*)</(?:\w+:)?t>", p)) for p in paragraphs]
    return html.unescape("\n".join(lines).strip("\n"))


def shapes_from_drawing(xml: str, rels: dict | None = None,
                        read: Callable[[str], bytes] | None = None) -> tuple[list[dict[str, Any]], int, dict[str, bytes]]:
    """Return the exportable shapes of a drawing, the number of shapes that
    cannot be exported and the pictures the shapes use.

    ``rels``: ``{rId: openpyxl Relationship}`` of the drawing (targets
    resolved); ``read``: reads a part of the package.
    """
    root, children = drawing_children(xml)
    shapes = []
    unsupported = 0
    media: dict[str, bytes] = {}
    graphics = 0
    for child in children:
        if _is_graphic(child):
            graphics += 1
            continue
        kind = classify(child)
        if kind == "unsupported shape":
            unsupported += 1
        elif kind == "shape":
            shape_media: dict[str, bytes] = {}
            shape_rels = _shape_rels(child, rels or {}, read, shape_media)
            if shape_rels is None:
                unsupported += 1
                continue
            media.update(shape_media)
            fragment = _self_contained(child, root)
            shape: dict[str, Any] = {}
            name = re.search(r"<(?:\w+:)?cNvPr\b[^>]*\bname=\"([^\"]*)\"", fragment)
            if name:
                shape["name"] = html.unescape(name.group(1))
            text = _text(fragment)
            if text:
                shape["text"] = text
            shape["layer"] = graphics
            if shape_rels:
                shape["rels"] = shape_rels
            shape["xml"] = fragment
            shapes.append(shape)
    for shape in shapes:
        if shape["layer"] >= graphics:
            del shape["layer"]
    return shapes, unsupported, media


def shape_cell(shape: dict[str, Any]) -> str | None:
    """Top-left cell of a shape, for summaries."""
    match = re.search(r"<(?:\w+:)?from>.*?<(?:\w+:)?col>(\d+)<.*?<(?:\w+:)?row>(\d+)<", shape.get("xml", ""), re.S)
    if not match:
        return None
    from openpyxl.utils import get_column_letter
    return f"{get_column_letter(int(match.group(1)) + 1)}{int(match.group(2)) + 1}"


def _renumber(fragments: list[str], mapping: dict[str, str]) -> list[str]:
    def tag(match: re.Match) -> str:
        return _ID_ATTR.sub(lambda m: f'id="{mapping.get(m.group(1), m.group(1))}"', match.group(0))
    return [_ID_TAG.sub(tag, fragment) for fragment in fragments]


def _shape_ids(fragments: list[str]) -> list[str]:
    ids = []
    for fragment in fragments:
        for match in _ID_TAG.finditer(fragment):
            if match.group(1) == "cNvPr":
                found = _ID_ATTR.search(match.group(0))
                if found and found.group(1) not in ids:
                    ids.append(found.group(1))
    return ids


def normalized(shapes: list[dict[str, Any]]) -> list[str]:
    """Shape XML with ids numbered in order and relationships resolved, for
    comparison: ids and relationship ids change on import.
    """
    fragments = []
    for shape in shapes:
        rels = shape.get("rels") or {}

        def resolve(match: re.Match) -> str:
            rel = rels.get(match.group(2))
            value = (rel.get("file") or rel.get("target")) if rel else match.group(2)
            return f'r:{match.group(1)}="{html.escape(str(value))}"'
        fragment = _REL_ATTR.sub(resolve, shape.get("xml", ""))
        fragments.append(f"layer={shape.get('layer')}|{fragment}")
    # Refer to shapes by name: ids are renumbered on import.
    names: dict[str, str] = {}
    for fragment in fragments:
        for tag in re.findall(r"<(?:\w+:)?cNvPr\b[^>]*>", fragment):
            found = _ID_ATTR.search(tag)
            name = re.search(r"\bname=\"([^\"]*)\"", tag)
            if found and found.group(1) not in names:
                names[found.group(1)] = f"#{name.group(1) if name else len(names) + 1}"
    return _renumber(fragments, names)


def rewrite_rel_ids(fragment: str, mapping: dict[str, str]) -> str:
    return _REL_ATTR.sub(lambda m: f'r:{m.group(1)}="{mapping.get(m.group(2), m.group(2))}"', fragment)


def add_to_drawing(drawing_xml: str, shapes: list[tuple[str, int | None]]) -> str:
    """Insert ``(fragment, layer)`` shapes into a drawing, keeping shape ids unique.

    A shape goes before the picture or chart number ``layer`` (at the end
    when ``layer`` is None or beyond the last one).
    """
    fragments = [fragment for fragment, _ in shapes]
    used = {int(i) for i in re.findall(r"<(?:\w+:)?cNvPr\b[^>]*\bid=\"(\d+)\"", drawing_xml)}
    ids = _shape_ids(fragments)
    if used & {int(i) for i in ids}:
        start = max(used | {int(i) for i in ids}) + 1
        fragments = _renumber(fragments, {old: str(start + n) for n, old in enumerate(ids)})

    _, children = drawing_children(drawing_xml)
    graphic_starts = []
    position = 0
    for child in children:
        position = drawing_xml.index(child, position)
        if _is_graphic(child):
            graphic_starts.append(position)
        position += len(child)
    end = drawing_xml.rindex("</")
    inserts: dict[int, list[str]] = {}
    for fragment, (_, layer) in zip(fragments, shapes):
        at = graphic_starts[layer] if layer is not None and layer < len(graphic_starts) else end
        inserts.setdefault(at, []).append(fragment)
    for at in sorted(inserts, reverse=True):
        drawing_xml = drawing_xml[:at] + "".join(inserts[at]) + drawing_xml[at:]
    return drawing_xml


EMPTY_DRAWING = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" '
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"></xdr:wsDr>'
)
