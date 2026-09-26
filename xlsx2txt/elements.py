"""Elements of the worksheet and workbook XML that openpyxl drops.

openpyxl reads what it has a model for and silently leaves out the rest of
``xl/worksheets/*.xml`` and ``xl/workbook.xml``: switched-off error checks,
ranges with their own password, custom views, "read-only recommended" and
more. Those elements are kept as self-contained XML fragments (``xmlElements``
of a sheet or of the workbook) and inserted back in the order the schema
requires.

Elements that refer to other parts through relationships (a background
picture, pictures in the header or footer) cannot be kept this way; they are
reported in the warnings instead.
"""

import re

from xlsx2txt.xmlfrag import insert_child, local_name, self_contained, split_children

# ECMA-376 CT_Worksheet, in schema order.
SHEET_ORDER = [
    "sheetPr", "dimension", "sheetViews", "sheetFormatPr", "cols", "sheetData", "sheetCalcPr",
    "sheetProtection", "protectedRanges", "scenarios", "autoFilter", "sortState", "dataConsolidate",
    "customSheetViews", "mergeCells", "phoneticPr", "conditionalFormatting", "dataValidations", "hyperlinks",
    "printOptions", "pageMargins", "pageSetup", "headerFooter", "rowBreaks", "colBreaks", "customProperties",
    "cellWatches", "ignoredErrors", "smartTags", "drawing", "legacyDrawing", "legacyDrawingHF", "drawingHF",
    "picture", "oleObjects", "controls", "webPublishItems", "tableParts", "extLst",
]
# Worksheet elements openpyxl drops that stand on their own.
SHEET_KEPT = {
    "sheetCalcPr", "protectedRanges", "sortState", "dataConsolidate", "customSheetViews", "phoneticPr",
    "cellWatches", "ignoredErrors", "smartTags", "webPublishItems",
}
# Worksheet elements openpyxl drops that use relationships.
SHEET_LINKED = {
    "picture": "background picture",
    "legacyDrawingHF": "pictures in the header/footer",
    "drawingHF": "pictures in the header/footer",
    "customProperties": "custom sheet properties",
}

# ECMA-376 CT_Workbook, in schema order.
WORKBOOK_ORDER = [
    "fileVersion", "fileSharing", "workbookPr", "workbookProtection", "bookViews", "sheets", "functionGroups",
    "externalReferences", "definedNames", "calcPr", "oleSize", "customWorkbookViews", "pivotCaches",
    "smartTagPr", "smartTagTypes", "webPublishing", "fileRecoveryPr", "webPublishObjects", "extLst",
]
# Workbook elements openpyxl drops that stand on their own. fileVersion is
# left out on purpose: Excel rewrites it (build numbers) on every save.
WORKBOOK_KEPT = {
    "fileSharing", "functionGroups", "oleSize", "customWorkbookViews", "smartTagPr", "smartTagTypes",
    "webPublishing", "fileRecoveryPr", "webPublishObjects",
}

_REL_REF = re.compile(r"\sr:\w+=")


def read(xml: str, kept: set[str]) -> tuple[list[str], list[str]]:
    """Return ``(fragments to keep, names of elements that cannot be kept)``."""
    root, children = split_children(xml)
    fragments, lost = [], []
    for child in children:
        name = local_name(child)
        if name not in kept:
            continue
        if _REL_REF.search(child):  # e.g. a custom view with its own printer settings
            lost.append(name)
        else:
            fragments.append(self_contained(child, root))
    return fragments, lost


def linked(xml: str) -> list[str]:
    """Descriptions of worksheet elements that use relationships and are lost."""
    _, children = split_children(xml)
    return [SHEET_LINKED[local_name(c)] for c in children if local_name(c) in SHEET_LINKED]


def write(xml: str, fragments: list[str], order: list[str]) -> str:
    """Insert fragments where the schema puts them."""
    for fragment in fragments:
        name = local_name(fragment)
        after = set(order[order.index(name) + 1:]) if name in order else set()
        xml = insert_child(xml, fragment, after)
    return xml


def name(fragment: str) -> str:
    return local_name(fragment)
