"""Round-trip fidelity checked independently of xlsx2txt's own model.

`verify` compares exports, so something the exporter never reads (e.g. page
headers before they were supported) cannot show up there. These tests open
each example and its restored copy with openpyxl and compare what openpyxl
sees: every cell's value, type and formatting, plus sheet and workbook
settings.
"""

import warnings
from pathlib import Path

import pytest
from openpyxl import load_workbook
from openpyxl.worksheet.formula import ArrayFormula

from xlsx2txt import export_model
from xlsx2txt.importer import import_model

EXAMPLES = sorted((Path(__file__).resolve().parent.parent / "examples").glob("*.xlsx"))


def _side(side):
    return None if side is None or (side.style is None and side.color is None) else repr(side)


def _border(border):
    # An empty <left/> and a missing one mean the same: no line.
    sides = tuple(_side(getattr(border, name)) for name in ("left", "right", "top", "bottom", "diagonal"))
    return sides + (bool(border.diagonalUp), bool(border.diagonalDown))


def _value(value):
    if isinstance(value, ArrayFormula):
        return ("array", value.ref, value.text)
    return (type(value).__name__, str(value))


def _cell(cell):
    return {
        "value": _value(cell.value),
        "type": cell.data_type,
        "numFmt": cell.number_format,
        "font": repr(cell.font),
        "fill": repr(cell.fill),
        "border": _border(cell.border),
        "alignment": repr(cell.alignment),
        "protection": repr(cell.protection),
        "style": cell.style,
        "comment": cell.comment.text if cell.comment else None,
        "link": (cell.hyperlink.target, cell.hyperlink.location) if cell.hyperlink else None,
    }


def _sheet(ws):
    return {
        "merged": sorted(map(str, ws.merged_cells.ranges)),
        "conditionalFormatting": sum(len(cf.rules) for cf in ws.conditional_formatting),
        "dataValidations": len(ws.data_validations.dataValidation),
        "tables": sorted(ws.tables),
        "charts": len(ws._charts),
        "images": len(ws._images),
        "pivots": len(ws._pivots),
        "freeze": ws.freeze_panes,
        "filter": ws.auto_filter.ref,
        "names": sorted(ws.defined_names),
        "state": ws.sheet_state,
        "protection": repr(ws.protection),
        "hiddenColumns": sorted(k for k, v in ws.column_dimensions.items() if v.hidden),
        "hiddenRows": sorted(k for k, v in ws.row_dimensions.items() if v.hidden),
        "outlineRows": sorted(k for k, v in ws.row_dimensions.items() if v.outline_level),
        "printArea": ws.print_area,
        "printTitles": (ws.print_title_rows, ws.print_title_cols),
        "headerFooter": repr(ws.HeaderFooter.to_tree() is not None and
                             [(child.tag, child.text) for child in ws.HeaderFooter.to_tree()]),
        "breaks": ([b.id for b in ws.row_breaks.brk], [b.id for b in ws.col_breaks.brk]),
        "margins": repr(ws.page_margins),
        "pageSetup": (ws.page_setup.orientation, ws.page_setup.fitToWidth, ws.page_setup.paperSize),
        "printOptions": repr(ws.print_options),
        "view": (ws.sheet_view.zoomScale, ws.sheet_view.showGridLines, ws.sheet_view.rightToLeft),
        "tabColor": repr(ws.sheet_properties.tabColor),
    }


def _workbook(wb):
    return {
        "sheets": wb.sheetnames,
        "chartsheets": [cs.title for cs in wb.chartsheets],
        "names": sorted(wb.defined_names),
        "namedStyles": sorted(wb.named_styles),
        "properties": {name: getattr(wb.properties, name) for name in
                       ("title", "subject", "creator", "keywords", "category", "description")},
        "customProperties": [(p.name, str(p.value)) for p in wb.custom_doc_props],
        # An empty <workbookProtection/> (written by openpyxl) means no protection.
        "protection": {name: value for name, value in vars(wb.security).items() if value} if wb.security else {},
    }


@pytest.mark.parametrize("path", EXAMPLES, ids=lambda p: p.stem)
def test_example_fidelity(path, tmp_path):
    restored = tmp_path / path.name
    import_model(export_model(path), restored)
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")  # openpyxl's notes about extensions it drops itself
        original, copy = load_workbook(path, rich_text=True), load_workbook(restored, rich_text=True)

    assert _workbook(copy) == _workbook(original)
    for ws in original.worksheets:
        other = copy[ws.title]
        assert _sheet(other) == _sheet(ws), ws.title
        for row in ws.iter_rows():
            for cell in row:
                assert _cell(other[cell.coordinate]) == _cell(cell), f"{ws.title}!{cell.coordinate}"
