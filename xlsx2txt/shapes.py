"""Shapes (text boxes, arrows, rectangles, connectors, groups) of worksheets.

openpyxl reads only pictures and charts from drawings, so shapes are taken
from the drawing XML directly: each drawing anchor holding a shape is kept
as a self-contained XML fragment and written back into the drawing of the
saved file on import.
"""

import html
import re
from typing import Any

# Tags of the drawing XML (comments and processing instructions included so
# they can be skipped when tracking nesting).
_TAG = re.compile(r"<!--.*?-->|<\?.*?\?>|<!\[CDATA\[.*?\]\]>|<(/?)([\w.:-]+)((?:[^>\"']|\"[^\"]*\"|'[^']*')*?)(/?)>",
                  re.S)
_ANCHORS = {"twoCellAnchor", "oneCellAnchor", "absoluteAnchor", "AlternateContent"}
# The first graphic object in an anchor decides what it is.
_OBJECT = re.compile(r"<(?:\w+:)?(sp|grpSp|cxnSp|pic|graphicFrame|contentPart)\b")
_SHAPE_KINDS = {"sp", "grpSp", "cxnSp"}
# References to other parts (pictures used as fills, hyperlinks, form controls).
_EXTERNAL = re.compile(r"\br:(embed|link|id|pict)=|compatExt")
_XMLNS = re.compile(r"\bxmlns(?::([\w.-]+))?=(\"[^\"]*\"|'[^']*')")
_ID_TAG = re.compile(r"<(?:\w+:)?(cNvPr|stCxn|endCxn)\b[^>]*>")
_ID_ATTR = re.compile(r"\bid=\"(\d+)\"")


def _local(name: str) -> str:
    return name.rsplit(":", 1)[-1]


def drawing_children(xml: str) -> tuple[str, list[str]]:
    """Split a drawing into its root start tag and its top-level elements."""
    depth = 0
    root = ""
    start = None
    children = []
    for match in _TAG.finditer(xml):
        closing, name, _, self_closing = match.groups()
        if name is None:
            continue
        if closing:
            depth -= 1
            if depth == 1 and start is not None:
                children.append(xml[start:match.end()])
                start = None
        else:
            if depth == 0:
                root = match.group(0)
            elif depth == 1:
                start = match.start()
                if self_closing:
                    children.append(match.group(0))
                    start = None
            if not self_closing:
                depth += 1
    return root, children


def classify(fragment: str) -> str | None:
    """Return "shape", "linked shape" (refers to other parts) or None."""
    match = re.match(r"<([\w.:-]+)", fragment)
    if not match or _local(match.group(1)) not in _ANCHORS:
        return None
    obj = _OBJECT.search(fragment)
    if not obj or obj.group(1) not in _SHAPE_KINDS:
        return None
    return "linked shape" if _EXTERNAL.search(fragment) else "shape"


def _self_contained(fragment: str, root: str) -> str:
    """Declare on the fragment the namespaces it inherits from the drawing root."""
    first_tag = re.match(r"<[^>]*?(?=/?>)", fragment).group(0)
    declared = {prefix or "" for prefix, _ in _XMLNS.findall(first_tag)}
    used = set(re.findall(r"</?([\w.-]+):", fragment))
    used |= set(re.findall(r"\s([\w.-]+):[\w.-]+=", fragment)) - {"xmlns"}
    for value in re.findall(r"\bIgnorable=\"([^\"]*)\"", fragment):
        used |= set(value.split())
    if re.search(r"<(?![\w.-]+:)[\w.-]+[\s/>]", fragment):
        used.add("")
    additions = [
        f" xmlns:{prefix}={value}" if prefix else f" xmlns={value}"
        for prefix, value in _XMLNS.findall(root)
        if (prefix or "") in used - declared
    ]
    return first_tag + "".join(additions) + fragment[len(first_tag):]


def _text(fragment: str) -> str:
    paragraphs = re.split(r"</(?:\w+:)?p>", fragment)
    lines = ["".join(re.findall(r"<(?:\w+:)?t>([^<]*)</(?:\w+:)?t>", p)) for p in paragraphs]
    return html.unescape("\n".join(lines).strip("\n"))


def shapes_from_drawing(xml: str) -> tuple[list[dict[str, Any]], int]:
    """Return the exportable shapes of a drawing and the number of linked ones."""
    root, children = drawing_children(xml)
    shapes = []
    linked = 0
    for child in children:
        kind = classify(child)
        if kind == "linked shape":
            linked += 1
        elif kind == "shape":
            fragment = _self_contained(child, root)
            shape: dict[str, Any] = {}
            name = re.search(r"<(?:\w+:)?cNvPr\b[^>]*\bname=\"([^\"]*)\"", fragment)
            if name:
                shape["name"] = html.unescape(name.group(1))
            text = _text(fragment)
            if text:
                shape["text"] = text
            shape["xml"] = fragment
            shapes.append(shape)
    return shapes, linked


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
    """Shape XML with ids numbered in order, for comparison.

    Ids are renumbered on import when they clash with pictures and charts.
    """
    fragments = [shape.get("xml", "") for shape in shapes]
    ids = _shape_ids(fragments)
    return _renumber(fragments, {old: str(number) for number, old in enumerate(ids, start=1)})


def add_to_drawing(drawing_xml: str, fragments: list[str]) -> str:
    """Append shape fragments to a drawing, keeping shape ids unique."""
    used = {int(i) for i in re.findall(r"<(?:\w+:)?cNvPr\b[^>]*\bid=\"(\d+)\"", drawing_xml)}
    ids = _shape_ids(fragments)
    if used & {int(i) for i in ids}:
        start = max(used | {int(i) for i in ids}) + 1
        fragments = _renumber(fragments, {old: str(start + n) for n, old in enumerate(ids)})
    end = drawing_xml.rindex("</")
    return drawing_xml[:end] + "".join(fragments) + drawing_xml[end:]


EMPTY_DRAWING = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" '
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"></xdr:wsDr>'
)
