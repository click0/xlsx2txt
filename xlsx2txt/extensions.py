"""Worksheet extensions (``<extLst>``) that openpyxl drops.

Sparklines, extended conditional formatting (data bar options, custom icon
sets) and extended data validation (lists from other sheets) live in the
``<extLst>`` of a worksheet. Every ``<ext>`` that does not refer to other
parts is kept as a self-contained XML fragment and written back on import.

Conditional formatting rules that have extended options point at them with
an id (``<x14:id>`` in the rule's own ``<extLst>``); openpyxl drops that
link too, so it is exported with the rule as ``extId`` and restored.

Two more pieces of worksheet XML openpyxl drops are kept here as well:
``<ignoredErrors>`` (error checks the user switched off, the green
triangles) and the ``cm`` attribute of cells, which points dynamic array
formulas (FILTER, SORT, UNIQUE...) at ``xl/metadata.xml``; without it Excel
shows them as legacy ``{=...}`` array formulas.
"""

import html
import re
from typing import Any

from xlsx2txt.xmlfrag import insert_child, self_contained, split_children

X14_NS = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
CF_EXT_URI = "{B025F937-C7B1-47D3-B67F-A62EFF666E3E}"

KINDS = {
    "{05C60535-1F16-4fd2-B633-F4F36F0B64E0}": "sparklines",
    "{78C0D931-6437-407d-A8EE-F0AAD7539E65}": "conditional formatting",
    "{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}": "data validation",
    "{A8765BA9-456A-4dab-B4F3-ACF838C121DE}": "slicers",
    "{3A4CF648-6AED-40f4-86FF-DC5316D8AED3}": "slicers",
    "{7E03D99C-DC04-49d9-9315-930204A7B6E9}": "timelines",
    "{F7C9EE02-42E1-4005-9D12-6889AFFD525C}": "Office add-ins",
}
# A reference to a relationship of the sheet (slicers, timelines, add-ins).
_REL_REF = re.compile(r"\sr:\w+=")
_URI = re.compile(r"\buri=\"([^\"]*)\"")
_CF_BLOCK = re.compile(r"<(?:\w+:)?conditionalFormatting\b([^>]*)>(.*?)</(?:\w+:)?conditionalFormatting>", re.S)
_CF_RULE = re.compile(r"<((?:\w+:)?)cfRule\b([^>]*?)(?:/>|>(.*?)</\1cfRule>)", re.S)
_SQREF = re.compile(r"\bsqref=\"([^\"]*)\"")
_PRIORITY = re.compile(r"\bpriority=\"(\d+)\"")


def _local(tag: str) -> str:
    return re.match(r"<([\w.:-]+)", tag).group(1).rsplit(":", 1)[-1]


def kind(uri: str) -> str:
    return KINDS.get(uri, f"extension {uri}")


def cf_key(sqref: str, priority: int) -> str:
    """Identify a conditional formatting rule within a sheet."""
    return f"{' '.join(sorted(sqref.split()))}#{priority}"


def read_sheet(sheet_xml: str) -> tuple[list[dict[str, Any]], dict[str, str], list[str]]:
    """Return ``(extensions, {cf_key: extId}, [kinds that cannot be kept])``."""
    root, children = split_children(sheet_xml)
    extensions: list[dict[str, Any]] = []
    lost: list[str] = []
    for child in children:
        if _local(child) != "extLst":
            continue
        ext_root, exts = split_children(child)
        for ext in exts:
            uri = _URI.search(ext.split(">", 1)[0])
            uri = uri.group(1) if uri else ""
            if _REL_REF.search(ext):
                lost.append(kind(uri))
                continue
            extensions.append({"type": kind(uri), "xml": self_contained(ext, root + ext_root)})

    ids: dict[str, str] = {}
    for attrs, body in _CF_BLOCK.findall(sheet_xml):
        sqref = _SQREF.search(attrs)
        if not sqref:
            continue
        for rule in re.findall(r"<(?:\w+:)?cfRule\b[^>]*?[^/]>.*?</(?:\w+:)?cfRule>", body, re.S):
            priority = _PRIORITY.search(rule.split(">", 1)[0])
            ext_id = re.search(rf"uri=\"{re.escape(CF_EXT_URI)}\".*?<(?:\w+:)?id>([^<]+)</", rule, re.S)
            if priority and ext_id:
                ids[cf_key(sqref.group(1), int(priority.group(1)))] = html.unescape(ext_id.group(1))
    return extensions, ids, lost


def _link_rules(sheet_xml: str, ids: dict[str, str]) -> str:
    def block(match: re.Match) -> str:
        sqref = _SQREF.search(match.group(1))
        if not sqref:
            return match.group(0)

        def rule(rule_match: re.Match) -> str:
            prefix, attrs, inner = rule_match.groups()
            priority = _PRIORITY.search(attrs)
            ext_id = ids.get(cf_key(sqref.group(1), int(priority.group(1)))) if priority else None
            if not ext_id:
                return rule_match.group(0)
            # The rule's extLst is its last child.
            link = (f'<{prefix}extLst><{prefix}ext uri="{CF_EXT_URI}" xmlns:x14="{X14_NS}">'
                    f"<x14:id>{html.escape(ext_id)}</x14:id></{prefix}ext></{prefix}extLst>")
            return f"<{prefix}cfRule{attrs}>{inner or ''}{link}</{prefix}cfRule>"
        body = _CF_RULE.sub(rule, match.group(2))
        return match.group(0)[:match.start(2) - match.start()] + body + match.group(0)[match.end(2) - match.start():]
    return _CF_BLOCK.sub(block, sheet_xml)


def write_sheet(sheet_xml: str, extensions: list[dict[str, Any]], ids: dict[str, str]) -> str:
    """Add extensions and conditional formatting links to a sheet saved by openpyxl."""
    if ids:
        sheet_xml = _link_rules(sheet_xml, ids)
    if extensions:
        fragments = "".join(ext["xml"] for ext in extensions)
        existing = re.search(r"</(?:\w+:)?extLst>\s*</(?:\w+:)?worksheet>\s*$", sheet_xml)
        if existing:
            sheet_xml = sheet_xml[:existing.start()] + fragments + sheet_xml[existing.start():]
        else:
            end = sheet_xml.rindex("</")
            sheet_xml = sheet_xml[:end] + f"<extLst>{fragments}</extLst>" + sheet_xml[end:]
    return sheet_xml


# ---------------------------------------------------------------------------
# <ignoredErrors> and cell metadata
# ---------------------------------------------------------------------------

METADATA_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata"
METADATA_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"
# Elements that follow <ignoredErrors> in a worksheet (ECMA-376, CT_Worksheet).
_AFTER_IGNORED = {"smartTags", "drawing", "legacyDrawing", "legacyDrawingHF", "drawingHF", "picture",
                  "oleObjects", "controls", "webPublishItems", "tableParts", "extLst"}
_CELL_TAG = re.compile(r"<(?:\w+:)?c\b([^>]*?)/?>")
_REF = re.compile(r"\br=\"([A-Z]+[0-9]+)\"")
_CM = re.compile(r"\bcm=\"(\d+)\"")


def read_ignored_errors(sheet_xml: str) -> str | None:
    root, children = split_children(sheet_xml)
    for child in children:
        if _local(child) == "ignoredErrors":
            return self_contained(child, root)
    return None


def write_ignored_errors(sheet_xml: str, fragment: str) -> str:
    return insert_child(sheet_xml, fragment, _AFTER_IGNORED)


def read_cell_metadata(sheet_xml: str) -> dict[str, int]:
    """``{cell: cm}`` for cells that refer to cell metadata (dynamic arrays)."""
    result = {}
    for attrs in _CELL_TAG.findall(sheet_xml):
        cm = _CM.search(attrs)
        ref = _REF.search(attrs)
        if cm and ref:
            result[ref.group(1)] = int(cm.group(1))
    return result


def write_cell_metadata(sheet_xml: str, cells: dict[str, int]) -> str:
    def cell(match: re.Match) -> str:
        ref = _REF.search(match.group(1))
        cm = cells.get(ref.group(1)) if ref else None
        if cm is None or _CM.search(match.group(1)):
            return match.group(0)
        tag = match.group(0)
        end = len(tag) - (2 if tag.endswith("/>") else 1)
        return f'{tag[:end]} cm="{cm}"{tag[end:]}'
    return _CELL_TAG.sub(cell, sheet_xml)
