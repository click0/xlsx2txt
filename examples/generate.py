"""Generate the example workbooks and their xlsx2txt exports.

Run from the repository root:

    python examples/generate.py

Every example is built with openpyxl and saved with fixed document dates, so
running the script again produces the same content.
"""

import datetime
import io
import re
import sys
import zipfile
from pathlib import Path

from openpyxl import Workbook
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.chart import (
    AreaChart,
    BarChart,
    DoughnutChart,
    LineChart,
    PieChart,
    Reference,
    ScatterChart,
    Series,
)
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.comments import Comment
from openpyxl.drawing.image import Image
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.formatting.rule import (
    CellIsRule,
    ColorScaleRule,
    DataBarRule,
    FormulaRule,
    IconSetRule,
    Rule,
)
from openpyxl.packaging.custom import (
    BoolProperty,
    DateTimeProperty,
    FloatProperty,
    IntProperty,
    StringProperty,
)
from openpyxl.packaging.relationship import Relationship
from openpyxl.pivot.cache import CacheDefinition
from openpyxl.pivot.record import RecordList
from openpyxl.pivot.table import TableDefinition
from openpyxl.styles import (
    Alignment,
    Border,
    Color,
    Font,
    GradientFill,
    NamedStyle,
    PatternFill,
    Protection,
    Side,
)
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.styles.numbers import NumberFormat
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.workbook.protection import WorkbookProtection
from openpyxl.workbook.external_link.external import (
    ExternalBook,
    ExternalCell,
    ExternalLink,
    ExternalRow,
    ExternalSheetData,
    ExternalSheetDataSet,
    ExternalSheetNames,
)
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.filters import ColorFilter, FilterColumn, Filters, SortCondition, SortState
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.worksheet.page import PageMargins
from openpyxl.worksheet.pagebreak import Break
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.xml.functions import fromstring

EXAMPLES_DIR = Path(__file__).resolve().parent
FIXED_DATE = datetime.datetime(2026, 1, 1, 12, 0, 0)

PRODUCTS = [
    ("Widget", "Hardware", 120, 9.99, datetime.date(2026, 1, 5), True),
    ("Gadget", "Hardware", 45, 24.50, datetime.date(2026, 1, 12), True),
    ("Gizmo", "Hardware", 80, 14.25, datetime.date(2026, 2, 3), False),
    ("Support plan", "Services", 12, 199.00, datetime.date(2026, 2, 20), True),
    ("Training", "Services", 5, 450.00, datetime.date(2026, 3, 9), False),
    ("Консультація", "Послуги", 3, 1200.00, datetime.date(2026, 3, 30), True),
]
HEADER_FONT = Font(bold=True, color="FFFFFFFF")
HEADER_FILL = PatternFill("solid", fgColor="FF305496")


def _new_workbook(title: str) -> Workbook:
    wb = Workbook()
    wb.properties.title = title
    wb.properties.creator = "xlsx2txt examples"
    wb.properties.created = FIXED_DATE
    return wb


def _header(ws, values) -> None:
    ws.append(values)
    for cell in ws[ws.max_row]:
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center")


def _png(color, size) -> bytes:
    from PIL import Image as PILImage

    buffer = io.BytesIO()
    PILImage.new("RGB", size, color).save(buffer, format="PNG")
    return buffer.getvalue()


# ---------------------------------------------------------------------------
# Examples
# ---------------------------------------------------------------------------

def basic_data() -> tuple[Workbook, dict]:
    """Values of every type on a larger sheet: numbers of all sizes, dates,
    times, durations, booleans, errors, text in several scripts; column
    widths, hidden and grouped rows and columns, frozen panes, a filter."""
    wb = _new_workbook("Basic data")
    ws = wb.active
    ws.title = "Sales"
    _header(ws, ["Product", "Category", "Quantity", "Price", "Date", "In stock", "Code", "Note"])
    regions = ["North", "South", "East", "West"]
    for index in range(40):
        name, category, quantity, price, date, in_stock = PRODUCTS[index % len(PRODUCTS)]
        ws.append([
            f"{name} {index + 1:02d}", category, quantity * (index % 7 + 1), round(price * (1 + index / 50), 2),
            date + datetime.timedelta(days=index * 3), in_stock if index % 5 else not in_stock,
            f"{index * 37 % 1000:05d}",  # leading zeros: kept as text
            f"{regions[index % 4]} region" if index % 3 else None,
        ])
    for row in ws.iter_rows(min_row=2, min_col=4, max_col=4):
        row[0].number_format = "#,##0.00"
    for row in ws.iter_rows(min_row=2, min_col=5, max_col=5):
        row[0].number_format = "yyyy-mm-dd"
    for column, width in zip("ABCDEFGH", (18, 12, 10, 10, 12, 10, 8, 16)):
        ws.column_dimensions[column].width = width
    ws.column_dimensions["G"].hidden = True
    # Rows 12-21 grouped (collapsible), row 25 hidden.
    for row in range(12, 22):
        ws.row_dimensions[row].outline_level = 1
    ws.row_dimensions[25].hidden = True
    ws.freeze_panes = "B2"
    # A filter with criteria (two categories, a colour filter on the code)
    # and a sort by price, descending; the matching rows are hidden as Excel does.
    ws.auto_filter.ref = f"A1:H{ws.max_row}"
    ws.auto_filter.filterColumn.append(FilterColumn(colId=1, filters=Filters(filter=["Hardware", "Services"])))
    marked = wb._differential_styles.add(DifferentialStyle(fill=PatternFill("solid", bgColor="FFFFEB9C")))
    ws.auto_filter.filterColumn.append(FilterColumn(colId=6, colorFilter=ColorFilter(dxfId=marked)))
    ws.auto_filter.sortState = SortState(ref=f"A2:H{ws.max_row}", sortCondition=[
        SortCondition(ref=f"D2:D{ws.max_row}", descending=True)])
    for row in range(2, ws.max_row + 1):
        if ws.cell(row, 2).value not in ("Hardware", "Services"):
            ws.row_dimensions[row].hidden = True
    ws.column_dimensions["I"].hidden = True
    ws.column_dimensions["I"].width = 0  # hidden by dragging it to zero width

    values = wb.create_sheet("Values")
    _header(values, ["Kind", "Value", "Format"])
    samples = [
        ("Integer", 42, None), ("Negative", -1234567, "#,##0;[Red]-#,##0"), ("Large", 123456789012345, "0"),
        ("Small", 0.000001234, "0.00E+00"), ("Fraction", 0.75, "# ?/?"), ("Percent", 0.1234, "0.00%"),
        ("Currency", 1999.5, '#,##0.00 "₴"'), ("Accounting", -45.1, '_-* #,##0.00_-;-* #,##0.00_-;_-* "-"??_-;_-@_-'),
        ("Date", datetime.date(2026, 2, 28), "dd.mm.yyyy"),
        ("Date and time", datetime.datetime(2026, 3, 1, 14, 30, 15), "yyyy-mm-dd hh:mm:ss"),
        ("Time", datetime.time(8, 45), "hh:mm"), ("Duration", datetime.timedelta(hours=26, minutes=5), "[h]:mm"),
        ("True", True, None), ("False", False, None), ("Error", "#N/A", None),
        ("Text as number", "00123", "@"), ("Multi-line", "First line\nSecond line", None),
        ("Українська", "Привіт, світе", None), ("日本語", "こんにちは", None), ("Emoji", "Ready ✅ 🚀", None),
        ("Long text", "Lorem ipsum dolor sit amet, " * 8, None),
    ]
    for kind, value, number_format in samples:
        values.append([kind, value, number_format or "General"])
        cell = values.cell(values.max_row, 2)
        if number_format:
            cell.number_format = number_format
        if kind == "Error":
            cell.data_type = "e"
        if kind == "Multi-line":
            cell.alignment = Alignment(wrap_text=True)
    values.column_dimensions["A"].width = 16
    values.column_dimensions["B"].width = 30
    values.column_dimensions["C"].width = 30
    return wb, {"zero_width_columns": {"Sales": [9]}}  # openpyxl cannot write width 0


def styles() -> Workbook:
    """Fonts, fills (solid, pattern, gradient), borders of every style,
    alignment, protection, number formats, a named style, merged cells,
    conditional formatting and data validation of many kinds."""
    wb = _new_workbook("Styles")
    ws = wb.active
    ws.title = "Styles"

    ws.merge_cells("A1:E1")
    ws["A1"] = "Quarterly report"
    ws["A1"].font = Font(name="Arial", size=16, bold=True, color="FF1F4E79")
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    thin = Side(style="thin", color="FF999999")
    _header(ws, ["Quarter", "Revenue", "Costs", "Margin", "Status"])
    rows = [("Q1", 125000, 98000), ("Q2", 142000, 101000), ("Q3", 98000, 105000), ("Q4", 171000, 120000)]
    for quarter, revenue, costs in rows:
        ws.append([quarter, revenue, costs, (revenue - costs) / revenue, "open"])
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=5):
        for cell in row:
            cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)
        row[1].number_format = '#,##0 "UAH"'
        row[2].number_format = '#,##0 "UAH"'
        row[3].number_format = "0.0%"

    ws["A8"] = "Wrapped text that does not fit into one line of the cell"
    ws["A8"].alignment = Alignment(wrap_text=True, vertical="top")
    ws["B8"] = "Rotated"
    ws["B8"].alignment = Alignment(text_rotation=45)
    ws["C8"] = "Italic strike"
    ws["C8"].font = Font(italic=True, strike=True)
    ws["D8"] = "Theme color"
    ws["D8"].fill = PatternFill("solid", fgColor="FFFFF2CC")

    # A gallery of fonts, fills, borders and alignment.
    gallery = [
        ("Double underline", Font(underline="double"), None, None, None),
        ("Subscript", Font(vertAlign="subscript"), None, None, None),
        ("Theme color + tint", Font(color=Color(theme=4, tint=-0.25), bold=True), None, None, None),
        ("Outline, shadow", Font(outline=True, shadow=True, size=12), None, None, None),
        ("Pattern fill", None, PatternFill("darkGrid", fgColor="FF9BC2E6", bgColor="FFFFFFFF"), None, None),
        ("Gradient fill", None, GradientFill(stop=("FFFFFFFF", "FF4472C4"), degree=90), None, None),
        ("Path gradient", None, GradientFill(type="path", left=0.5, right=0.5, top=0.5, bottom=0.5,
                                             stop=("FFFFFFFF", "FFED7D31")), None, None),
        ("Double/thick borders", None, None, Border(left=Side("double"), right=Side("thick", color="FFC00000"),
                                                    top=Side("dashDot"), bottom=Side("mediumDashed")), None),
        ("Diagonal border", None, None, Border(diagonal=Side("thin"), diagonalUp=True, diagonalDown=True), None),
        ("Indented, shrink to fit", None, None, None, Alignment(indent=2, shrink_to_fit=True)),
        ("Justified, bottom", None, None, None, Alignment(horizontal="justify", vertical="bottom")),
        ("Vertical text", None, None, None, Alignment(text_rotation=255)),
    ]
    for offset, (text, font, fill, border, alignment) in enumerate(gallery):
        cell = ws.cell(10 + offset, 1, text)
        for attr, value in (("font", font), ("fill", fill), ("border", border), ("alignment", alignment)):
            if value is not None:
                setattr(cell, attr, value)

    formats = [
        ("Date", datetime.date(2026, 4, 1), "d mmmm yyyy"), ("Scientific", 6.02214076e23, "0.000E+00"),
        ("Fraction", 1.375, "# ??/??"), ("Text", 12345, "@"), ("Thousands", 1234567.891, "#,##0.0,\" K\""),
        ("Colored sign", -12.5, "[Green]0.0;[Red]-0.0;[Blue]0"), ("Leading zeros", 42, "000000"),
    ]
    for offset, (label, value, number_format) in enumerate(formats):
        ws.cell(10 + offset, 3, label)
        cell = ws.cell(10 + offset, 4, value)
        cell.number_format = number_format

    # Cells that stay editable on a protected sheet, and a hidden formula.
    ws["F3"] = "Editable"
    ws["F3"].protection = Protection(locked=False)
    ws["F4"] = "=B3-C3"
    ws["F4"].protection = Protection(hidden=True)

    highlight = NamedStyle(name="Highlight", font=Font(bold=True, color="FF7030A0"),
                           fill=PatternFill("solid", fgColor="FFE4DFEC"), border=Border(bottom=thin))
    wb.add_named_style(highlight)
    ws["F6"] = "Named style"
    ws["F6"].style = "Highlight"

    ws.conditional_formatting.add("D3:D6", CellIsRule(operator="lessThan", formula=["0"],
                                                       font=Font(color="FF9C0006"),
                                                       fill=PatternFill("solid", bgColor="FFFFC7CE")))
    ws.conditional_formatting.add("B3:B6", DataBarRule(start_type="min", end_type="max", color="FF638EC6"))
    ws.conditional_formatting.add("C3:C6", ColorScaleRule(start_type="min", start_color="FF63BE7B",
                                                          mid_type="percentile", mid_value=50, mid_color="FFFFEB84",
                                                          end_type="max", end_color="FFF8696B"))
    ws.conditional_formatting.add("B3:C6", IconSetRule("3Arrows", "percent", [0, 33, 67]))
    ws.conditional_formatting.add("A3:A6", FormulaRule(formula=['$D3<0.1'], font=Font(bold=True, color="FFC00000")))
    ws.conditional_formatting.add("B3:B6", Rule(type="top10", rank=1, dxf=DifferentialStyle(
        fill=PatternFill("solid", bgColor="FFC6EFCE"))))
    ws.conditional_formatting.add("E3:E6", Rule(type="duplicateValues", dxf=DifferentialStyle(
        font=Font(italic=True))))

    status = DataValidation(type="list", formula1='"open,closed,archived"', allow_blank=False,
                            showErrorMessage=True, error="Choose a status from the list")
    status.add("E3:E6")
    ws.add_data_validation(status)
    whole = DataValidation(type="whole", operator="between", formula1="0", formula2="1000000",
                           showInputMessage=True, promptTitle="Revenue", prompt="Whole number up to 1 000 000")
    whole.add("B3:C6")
    ws.add_data_validation(whole)
    date = DataValidation(type="date", operator="greaterThan", formula1="DATE(2026,1,1)")
    date.add("D10")
    ws.add_data_validation(date)
    length = DataValidation(type="textLength", operator="lessThanOrEqual", formula1="20", errorStyle="warning")
    length.add("F3")
    ws.add_data_validation(length)
    custom = DataValidation(type="custom", formula1="=MOD(F5,5)=0", error="Multiples of 5 only",
                            showErrorMessage=True)
    custom.add("F5")
    ws.add_data_validation(custom)

    for column, width in zip("ABCDEF", (24, 16, 16, 14, 12, 14)):
        ws.column_dimensions[column].width = width
    ws.protection.sheet = True
    ws.protection.formatCells = False  # formatting stays allowed
    return wb


def formulas() -> Workbook:
    """Formulas of many kinds: arithmetic, lookups, conditional sums, text and
    date functions, cross-sheet and 3D references, an array formula, named
    ranges and constants (workbook and sheet scope), tables with totals and
    structured references, a hidden sheet, comments and hyperlinks."""
    wb = _new_workbook("Formulas")
    ws = wb.active
    ws.title = "Orders"
    _header(ws, ["Product", "Quantity", "Price", "Total", "Category"])
    for index, (name, category, quantity, price, _, _) in enumerate(PRODUCTS, start=2):
        ws.append([name, quantity, price, f"=B{index}*C{index}", category])
    last = ws.max_row
    ws[f"A{last + 1}"] = "Sum"
    ws[f"D{last + 1}"] = f"=SUM(D2:D{last})"
    ws[f"A{last + 2}"] = "Average price"
    ws[f"D{last + 2}"] = f"=AVERAGE(C2:C{last})"
    ws[f"A{last + 3}"] = "With VAT"
    ws[f"D{last + 3}"] = f"=D{last + 1}*(1+VAT)"
    ws[f"A{last + 4}"] = "Hardware total"
    ws[f"D{last + 4}"] = f'=SUMIFS(D2:D{last},E2:E{last},"Hardware")'
    ws[f"A{last + 5}"] = "Orders over 1000"
    ws[f"D{last + 5}"] = f'=COUNTIF(D2:D{last},">1000")'
    ws["G1"] = "Doubled"
    ws["G2"] = ArrayFormula(f"G2:G{last}", f"=B2:B{last}*2")

    table = Table(displayName="OrderList", ref=f"A1:E{last}")
    table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    # Table formatting of its own (differential formats): a header font and a
    # number format for the Total column.
    table.headerRowDxfId = wb._differential_styles.add(DifferentialStyle(font=Font(b=True, color="FFFFFFFF")))
    table._initialise_columns()
    for column, cell in zip(table.tableColumns, ws[1]):
        column.name = cell.value  # column names must match the header cells
        if column.name == "Total":
            column.dataDxfId = wb._differential_styles.add(DifferentialStyle(
                numFmt=NumberFormat(numFmtId=164, formatCode="#,##0.00 [$UAH]")))
    ws.add_table(table)

    ws["A1"].comment = Comment("Product names come from the catalog", "Analyst")
    ws["D1"].comment = Comment("Total = Quantity × Price\nVAT is added in the summary", "Finance", 120, 240)
    ws[f"A{last + 7}"] = "Documentation"
    ws[f"A{last + 7}"].hyperlink = "https://github.com/click0/xlsx2txt"
    ws[f"A{last + 7}"].style = "Hyperlink"
    ws[f"A{last + 8}"] = "Go to the summary"
    ws[f"A{last + 8}"].hyperlink = Hyperlink(ref=f"A{last + 8}", location="Summary!A1", tooltip="Summary sheet")
    ws[f"A{last + 8}"].style = "Hyperlink"

    lookup = wb.create_sheet("Lookup")
    _header(lookup, ["Product", "Found price", "Row", "Label", "Due"])
    for row, (name, *_rest) in enumerate(PRODUCTS[:4], start=2):
        lookup.append([
            name,
            f"=VLOOKUP(A{row},Orders!$A$2:$C${last},3,FALSE)",
            f"=MATCH(A{row},Orders!$A$2:$A${last},0)",
            f'=UPPER(LEFT(A{row},3))&"-"&TEXT(B{row},"0.00")',
            f"=EDATE(DATE(2026,1,31),ROW()-1)",
        ])
    lookup["A7"] = "Safe division"
    lookup["B7"] = '=IFERROR(1/0,"n/a")'
    lookup["A8"] = "Index/match"
    lookup["B8"] = f'=INDEX(Orders!$D$2:$D${last},MATCH("Gadget",Orders!$A$2:$A${last},0))'
    lookup["A9"] = "Discount (sheet name)"
    lookup["B9"] = "=Discount"
    lookup.defined_names["Discount"] = DefinedName("Discount", attr_text="0.05")

    # Three months with the same layout for a 3D reference.
    for month, amount in (("Jan", 100), ("Feb", 150), ("Mar", 125)):
        sheet = wb.create_sheet(month)
        sheet["A1"], sheet["B1"] = "Amount", amount
    budget = wb.create_sheet("Budget")
    budget["A1"] = "Quarter"
    budget["B1"] = "=SUM(Jan:Mar!B1)"
    _header_row = ["Item", "Plan", "Actual", "Delta"]
    for column, value in enumerate(_header_row, start=1):
        budget.cell(3, column, value)
    for row, (item, plan, actual) in enumerate([("Rent", 1000, 1000), ("Staff", 5000, 5200),
                                                 ("Travel", 800, 650)], start=4):
        budget.cell(row, 1, item)
        budget.cell(row, 2, plan)
        budget.cell(row, 3, actual)
        budget.cell(row, 4, f"=BudgetTable[[#This Row],[Actual]]-BudgetTable[[#This Row],[Plan]]")
    budget["A7"] = "Total"
    budget["B7"] = "=SUBTOTAL(109,BudgetTable[Plan])"
    budget["C7"] = "=SUBTOTAL(109,BudgetTable[Actual])"
    budget["D7"] = "=SUBTOTAL(109,BudgetTable[Delta])"
    totals = Table(displayName="BudgetTable", ref="A3:D7", totalsRowCount=1)
    totals.tableStyleInfo = TableStyleInfo(name="TableStyleLight9", showRowStripes=True)
    totals._initialise_columns()
    for column, name, function in zip(totals.tableColumns, _header_row, (None, "sum", "sum", "sum")):
        column.name = name  # column names must match the header cells
        if function:
            column.totalsRowFunction = function
        else:
            column.totalsRowLabel = "Total"
    budget.add_table(totals)

    settings = wb.create_sheet("Settings")
    settings["A1"] = "VAT"
    settings["B1"] = 0.2
    settings.sheet_state = "hidden"
    wb.defined_names["VAT"] = DefinedName("VAT", attr_text="Settings!$B$1")
    wb.defined_names["OrderProducts"] = DefinedName("OrderProducts", attr_text=f"Orders!$A$2:$A${last}")

    summary = wb.create_sheet("Summary")
    summary["A1"] = "Orders total"
    summary["B1"] = f"=Orders!D{last + 1}"
    summary["A2"] = "Largest order"
    summary["B2"] = f"=MAX(Orders!D2:D{last})"
    summary["A3"] = "Products"
    summary["B3"] = "=COUNTA(OrderProducts)"
    summary["A4"] = "Table total"
    summary["B4"] = "=SUM(OrderList[Total])"
    summary["A5"] = "Quarter budget"
    summary["B5"] = "=Budget!B1"
    return wb


def images_and_charts() -> Workbook:
    """Images (PNG and JPEG, different anchors) and charts of many types:
    clustered and stacked bars, lines with markers, area, scatter, doughnut,
    a bar/line combination with a secondary axis; a chart sheet."""
    wb = _new_workbook("Images and charts")
    ws = wb.active
    ws.title = "Data"
    _header(ws, ["Month", "Revenue", "Costs", "Visitors", "Conversion"])
    months = [("Jan", 42, 30, 1200, 0.021), ("Feb", 47, 31, 1350, 0.024), ("Mar", 51, 35, 1600, 0.022),
              ("Apr", 49, 33, 1550, 0.026), ("May", 58, 36, 1800, 0.027), ("Jun", 63, 40, 2100, 0.025)]
    for row in months:
        ws.append(list(row))
    last = ws.max_row

    ws.add_image(Image(io.BytesIO(_png((48, 84, 150), (120, 40)))), "H1")
    jpeg = io.BytesIO()
    from PIL import Image as PILImage
    PILImage.new("RGB", (60, 60), (237, 125, 49)).save(jpeg, format="JPEG", quality=90)
    photo = Image(io.BytesIO(jpeg.getvalue()))
    photo.anchor = OneCellAnchor(_from=AnchorMarker(col=10, row=0, colOff=95250, rowOff=0),
                                 ext=XDRPositiveSize2D(cx=571500, cy=571500))
    ws.add_image(photo)

    categories = Reference(ws, min_col=1, min_row=2, max_row=last)
    bar = BarChart()
    bar.title = "Revenue and costs"
    bar.y_axis.title = "k UAH"
    bar.add_data(Reference(ws, min_col=2, max_col=3, min_row=1, max_row=last), titles_from_data=True)
    bar.set_categories(categories)
    ws.add_chart(bar, "H4")

    stacked = BarChart()
    stacked.type = "bar"
    stacked.grouping = "stacked"
    stacked.overlap = 100
    stacked.title = "Stacked (horizontal)"
    stacked.add_data(Reference(ws, min_col=2, max_col=3, min_row=1, max_row=last), titles_from_data=True)
    stacked.set_categories(categories)
    ws.add_chart(stacked, "P4")

    line = LineChart()
    line.title = "Revenue trend"
    line.add_data(Reference(ws, min_col=2, min_row=1, max_row=last), titles_from_data=True)
    line.set_categories(categories)
    line.series[0].marker.symbol = "circle"
    line.series[0].smooth = True
    ws.add_chart(line, "H20")

    area = AreaChart()
    area.title = "Costs (area)"
    area.add_data(Reference(ws, min_col=3, min_row=1, max_row=last), titles_from_data=True)
    area.set_categories(categories)
    ws.add_chart(area, "P20")

    scatter = ScatterChart()
    scatter.title = "Visitors vs revenue"
    scatter.style = 13
    scatter.x_axis.title = "Visitors"
    scatter.y_axis.title = "Revenue"
    series = Series(Reference(ws, min_col=2, min_row=2, max_row=last),
                    Reference(ws, min_col=4, min_row=2, max_row=last), title="Months")
    series.marker.symbol = "diamond"
    series.marker.size = 8
    series.marker.graphicalProperties = GraphicalProperties(solidFill="4472C4")
    series.graphicalProperties.line.noFill = True
    scatter.series.append(series)
    ws.add_chart(scatter, "H36")

    combo = BarChart()
    combo.title = "Visitors and conversion"
    combo.add_data(Reference(ws, min_col=4, min_row=1, max_row=last), titles_from_data=True)
    combo.set_categories(categories)
    rate = LineChart()
    rate.add_data(Reference(ws, min_col=5, min_row=1, max_row=last), titles_from_data=True)
    rate.y_axis.axId = 200
    rate.y_axis.title = "Conversion"
    rate.y_axis.crosses = "max"
    combo += rate
    ws.add_chart(combo, "P36")

    for chart in (bar, stacked, line, area, scatter, combo):
        chart.roundedCorners = False  # as Excel writes it

    pie = PieChart()
    pie.title = "Revenue share"
    pie.add_data(Reference(ws, min_col=2, min_row=1, max_row=last), titles_from_data=True)
    pie.set_categories(categories)
    wb.create_chartsheet("Pie chart").add_chart(pie)

    doughnut = DoughnutChart()
    doughnut.title = "Costs share"
    doughnut.add_data(Reference(ws, min_col=3, min_row=1, max_row=last), titles_from_data=True)
    doughnut.set_categories(categories)
    report = wb.create_sheet("Report")
    report["A1"] = "A chart on another sheet, built from the Data sheet"
    report.add_chart(doughnut, "A3")
    return wb


def rich_text() -> Workbook:
    """Cells with several differently formatted runs of text: bold, italic,
    underline, strike, colors, fonts and sizes, super- and subscript, line
    breaks, several scripts; rich text in merged cells and next to formulas."""
    wb = _new_workbook("Rich text")
    ws = wb.active
    ws.title = "Notes"
    ws["A1"] = CellRichText([
        "Status: ",
        TextBlock(InlineFont(b=True, color="FF00B050"), "approved"),
        " by ",
        TextBlock(InlineFont(i=True), "finance"),
    ])
    ws["A2"] = CellRichText([
        TextBlock(InlineFont(rFont="Courier New", sz=10), "E=mc"),
        TextBlock(InlineFont(rFont="Courier New", sz=10, vertAlign="superscript"), "2"),
    ])
    ws["A3"] = CellRichText([
        "Увага: ",
        TextBlock(InlineFont(b=True, u="single", color="FFC00000"), "термін до 31.03"),
    ])
    ws["A4"] = CellRichText([
        "H", TextBlock(InlineFont(vertAlign="subscript"), "2"), "O and CO",
        TextBlock(InlineFont(vertAlign="subscript"), "2"),
    ])
    ws["A5"] = CellRichText([
        TextBlock(InlineFont(strike=True, color="FF7F7F7F"), "old price 120"), " ",
        TextBlock(InlineFont(b=True, sz=14, color="FFC00000"), "99"), " UAH",
    ])
    ws["A6"] = CellRichText([
        TextBlock(InlineFont(b=True), "Line one"), "\n",
        TextBlock(InlineFont(i=True, color="FF2F5597"), "line two"), "\n",
        TextBlock(InlineFont(u="double"), "line three"),
    ])
    ws["A6"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[6].height = 48
    ws["A7"] = CellRichText([
        TextBlock(InlineFont(rFont="Georgia", sz=12), "Georgia "),
        TextBlock(InlineFont(rFont="Verdana", sz=9), "Verdana "),
        TextBlock(InlineFont(rFont="Times New Roman", sz=11, i=True), "Times"),
    ])
    ws["A8"] = CellRichText([
        "日本語 ", TextBlock(InlineFont(b=True), "太字"), " · Ελληνικά ",
        TextBlock(InlineFont(i=True), "πλάγια"), " · ✅",
    ])
    ws.merge_cells("A10:C11")
    ws["A10"] = CellRichText([
        TextBlock(InlineFont(b=True, sz=13, color="FF1F4E79"), "Merged: "),
        "rich text spanning three columns and two rows",
    ])
    ws["A10"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws["A13"] = "Plain text next to a formula"
    ws["B13"] = '=LEN(A13)'
    ws.column_dimensions["A"].width = 40
    ws.column_dimensions["B"].width = 12
    return wb


def workbook_features() -> tuple[Workbook, dict]:
    """Link to another workbook, document and custom properties of every type,
    sheet and workbook protection, printing (area, titles, headers/footers,
    page breaks, margins), views (zoom, gridlines, right-to-left), hidden and
    very hidden sheets, a sheet-scoped name, an editable range with its own
    password, a custom view and "open as read-only" recommended."""
    wb = _new_workbook("Workbook features")
    ws = wb.active
    ws.title = "Report"
    ws["A1"] = "Budget from another file"
    ws["B1"] = "=[1]Budget!B2"
    ws["A2"] = "Protected sheet, print settings: landscape, fit to width"
    ws["A3"] = "Cell B3 can be edited with its own password (range 'Input')"
    ws["B3"] = 0

    book = ExternalBook(
        sheetNames=ExternalSheetNames(sheetName=["Budget"]),
        sheetDataSet=ExternalSheetDataSet(sheetData=[
            ExternalSheetData(sheetId=0, row=[ExternalRow(r=2, cell=[ExternalCell(r="B2", v="250000")])]),
        ]),
        id="rId1",
    )
    link = ExternalLink(externalBook=book)
    link.file_link = Relationship(type="externalLinkPath", Target="budget-2026.xlsx", TargetMode="External")
    wb._external_links.append(link)

    wb.custom_doc_props.append(StringProperty(name="Department", value="Finance"))
    wb.custom_doc_props.append(BoolProperty(name="Approved", value=True))
    wb.custom_doc_props.append(IntProperty(name="Revision", value=7))
    wb.custom_doc_props.append(FloatProperty(name="Budget share", value=0.35))
    wb.custom_doc_props.append(DateTimeProperty(name="Review date", value=FIXED_DATE))
    wb.properties.subject = "Budget 2026"
    wb.properties.keywords = "budget; finance; example"
    wb.properties.category = "Reports"
    wb.properties.description = "Shows workbook-level features of xlsx2txt"
    wb.security = WorkbookProtection(lockStructure=True)

    for row in range(5, 45):
        ws.cell(row, 1, f"Line {row - 4}")
        ws.cell(row, 2, (row * 37) % 101)
    ws.print_area = "A1:B44"
    ws.row_breaks.append(Break(id=24))
    ws.col_breaks.append(Break(id=1))
    ws.page_margins = PageMargins(left=0.5, right=0.5, top=0.8, bottom=0.8, header=0.3, footer=0.3)
    ws.oddHeader.center.text = "&[File] — &[Tab]"
    ws.oddHeader.right.text = "&D"
    ws.oddFooter.center.text = "Page &[Page] of &N"
    ws.oddFooter.left.text = "Confidential"
    ws.print_options.gridLines = True
    ws.print_options.horizontalCentered = True
    ws.sheet_view.zoomScale = 90
    ws.sheet_view.showGridLines = False

    ws.protection.sheet = True
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = "1:1"
    ws.sheet_properties.tabColor = "FFC00000"
    ws.column_dimensions["A"].width = 60
    ws.protection.formatColumns = False
    ws.protection.sort = False

    contacts = wb.create_sheet("Contacts")
    contacts.sheet_view.rightToLeft = True
    contacts.sheet_properties.tabColor = "FF00B050"
    contacts.append(["Name", "Phone"])
    contacts.append(["Olena", "+380 44 000 0000"])
    contacts.defined_names["FirstContact"] = DefinedName("FirstContact", attr_text="Contacts!$A$2")
    archive = wb.create_sheet("Archive")
    archive["A1"] = "Hidden sheet"
    archive.sheet_state = "hidden"
    internal = wb.create_sheet("Internal")
    internal["A1"] = "Very hidden: only VBA or the XML can show it"
    internal.sheet_state = "veryHidden"
    # openpyxl has no model for these; they are added to the saved file. The
    # password hash of the range is a placeholder, not a real password.
    return wb, {
        "sheet_elements": {"Report": [
            '<protectedRanges><protectedRange sqref="B3" name="Input" algorithmName="SHA-512" '
            'hashValue="Dg5gGu0VbAqKeUzRz6dHo9nJ0mkPzqtNUnvWIBNExmuYrfgjeRUvePjPH8zhL2CrYhhlymzCP9Pyo6EZAwN7Vw==" '
            'saltValue="yMNjgUq6n7Hs2cRldfQ+Bg==" spinCount="100000"/></protectedRanges>',
            '<customSheetViews><customSheetView guid="{3F2504E0-4F89-11D3-9A0C-0305E82C3301}" scale="120" '
            'showGridLines="0"><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" '
            'footer="0.3"/></customSheetView></customSheetViews>',
        ]},
        "workbook_elements": ['<fileSharing readOnlyRecommended="1"/>'],
    }


# openpyxl cannot create pivot tables, so this one is described in XML the way
# Excel stores it: cache definition, cached records and the table itself.
PIVOT_MAIN = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
PIVOT_REL = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'

PIVOT_DATA = [("Category", "Amount"), ("Fruit", 1), ("Veg", 2), ("Fruit", 3), ("Veg", 4)]

PIVOT_CACHE = f"""<pivotCacheDefinition {PIVOT_MAIN} {PIVOT_REL} refreshOnLoad="1" recordCount="4">
  <cacheSource type="worksheet"><worksheetSource ref="A1:B5" sheet="Data"/></cacheSource>
  <cacheFields count="2">
    <cacheField name="Category" numFmtId="0">
      <sharedItems count="2"><s v="Fruit"/><s v="Veg"/></sharedItems>
    </cacheField>
    <cacheField name="Amount" numFmtId="0">
      <sharedItems containsSemiMixedTypes="0" containsString="0" containsNumber="1"
                   containsInteger="1" minValue="1" maxValue="4"/>
    </cacheField>
  </cacheFields>
</pivotCacheDefinition>"""

PIVOT_RECORDS = f"""<pivotCacheRecords {PIVOT_MAIN} count="4">
  <r><x v="0"/><n v="1"/></r>
  <r><x v="1"/><n v="2"/></r>
  <r><x v="0"/><n v="3"/></r>
  <r><x v="1"/><n v="4"/></r>
</pivotCacheRecords>"""

PIVOT_TABLE = f"""<pivotTableDefinition {PIVOT_MAIN} name="Totals" cacheId="1" dataCaption="Values"
    applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0"
    applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="1"
    updatedVersion="6" minRefreshableVersion="3" createdVersion="6" indent="0"
    outline="1" outlineData="1">
  <location ref="D1:E4" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"/>
  <pivotFields count="2">
    <pivotField axis="axisRow" showAll="0">
      <items count="3"><item x="0"/><item x="1"/><item t="default"/></items>
    </pivotField>
    <pivotField dataField="1" showAll="0"/>
  </pivotFields>
  <rowFields count="1"><field x="0"/></rowFields>
  <rowItems count="3"><i><x/></i><i><x v="1"/></i><i t="grand"><x/></i></rowItems>
  <colItems count="1"><i/></colItems>
  <dataFields count="1"><dataField name="Sum of Amount" fld="1" baseField="0" baseItem="0"/></dataFields>
  <pivotTableStyleInfo name="PivotStyleLight16" showRowHeaders="1" showColHeaders="1"
      showRowStripes="0" showColStripes="0" showLastColumn="1"/>
</pivotTableDefinition>"""



PIVOT_COUNT = (PIVOT_TABLE.replace('name="Totals"', 'name="Counts"')
               .replace('<location ref="D1:E4"', '<location ref="A3:B6"')
               .replace('name="Sum of Amount" fld="1"', 'name="Count of Amount" fld="1" subtotal="count"'))


def pivot_table() -> Workbook:
    """Two pivot tables built on one cache: amounts summed by category next to
    the data, and counted by category on another sheet."""
    wb = _new_workbook("Pivot table")
    ws = wb.active
    ws.title = "Data"
    for row in PIVOT_DATA:
        ws.append(row)
    # The values of the pivot table as Excel shows (and stores) them.
    for coord, value in {"D1": "Row Labels", "E1": "Sum of Amount", "D2": "Fruit", "E2": 4,
                         "D3": "Veg", "E3": 6, "D4": "Grand Total", "E4": 10}.items():
        ws[coord] = value

    cache = CacheDefinition.from_tree(fromstring(PIVOT_CACHE))
    cache.records = RecordList.from_tree(fromstring(PIVOT_RECORDS))
    pivot = TableDefinition.from_tree(fromstring(PIVOT_TABLE))
    pivot.cache = cache
    ws.add_pivot(pivot)

    summary = wb.create_sheet("Summary")
    summary["A1"] = "Orders per category (a second pivot table on the same cache)"
    for coord, value in {"A3": "Row Labels", "B3": "Count of Amount", "A4": "Fruit", "B4": 2,
                         "A5": "Veg", "B5": 2, "A6": "Grand Total", "B6": 4}.items():
        summary[coord] = value
    counts = TableDefinition.from_tree(fromstring(PIVOT_COUNT))
    counts.cache = cache
    summary.add_pivot(counts)
    return wb


# ---------------------------------------------------------------------------
# Shapes: openpyxl cannot create them, so their drawing XML (as Excel writes
# it) is added to the saved file.
# ---------------------------------------------------------------------------

_DRAWING_NS = ('xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" '
               'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
               'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')


def _anchor(start, end, body):
    (col1, row1), (col2, row2) = start, end
    return (f'<xdr:twoCellAnchor {_DRAWING_NS}>'
            f'<xdr:from><xdr:col>{col1}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{row1}</xdr:row>'
            f'<xdr:rowOff>0</xdr:rowOff></xdr:from>'
            f'<xdr:to><xdr:col>{col2}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{row2}</xdr:row>'
            f'<xdr:rowOff>0</xdr:rowOff></xdr:to>{body}<xdr:clientData/></xdr:twoCellAnchor>')


def _shape(shape_id, name, geometry, fill, text=(), text_box=False, offset=(0, 0), size=(0, 0),
           link=None, picture=None):
    paragraphs = "".join(f'<a:p><a:r><a:rPr lang="en-US" sz="1100"/><a:t>{line}</a:t></a:r></a:p>'
                         for line in text) or '<a:p><a:endParaRPr lang="en-US" sz="1100"/></a:p>'
    text_box_attr = ' txBox="1"' if text_box else ""
    fill_xml = f'<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>' if fill else "<a:noFill/>"
    if picture:
        fill_xml = f'<a:blipFill><a:blip r:embed="{picture}"/><a:stretch><a:fillRect/></a:stretch></a:blipFill>'
    link_xml = f'<a:hlinkClick r:id="{link}"/>' if link else ""
    return (f'<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="{shape_id}" name="{name}">{link_xml}'
            f'</xdr:cNvPr>'
            f'<xdr:cNvSpPr{text_box_attr}/></xdr:nvSpPr>'
            f'<xdr:spPr><a:xfrm><a:off x="{offset[0]}" y="{offset[1]}"/><a:ext cx="{size[0]}" cy="{size[1]}"/>'
            f'</a:xfrm><a:prstGeom prst="{geometry}"><a:avLst/></a:prstGeom>{fill_xml}'
            f'<a:ln w="9525"><a:solidFill><a:srgbClr val="305496"/></a:solidFill></a:ln></xdr:spPr>'
            f'<xdr:txBody><a:bodyPr wrap="square" rtlCol="0" anchor="t"/><a:lstStyle/>{paragraphs}</xdr:txBody>'
            f'</xdr:sp>')


_SHAPES_XML = {
    "Diagram": [
        _anchor((0, 1), (2, 7), _shape(9, "Frame 8", "rect", "D9D9D9", text=("Behind the picture",))),
        _anchor((4, 1), (8, 6), _shape(2, "TextBox 1", "rect", None, text_box=True, text=(
            "Text box with two paragraphs.", "Shapes are kept as drawing XML."))),
        _anchor((1, 8), (3, 12), _shape(3, "Rectangle 2", "rect", "DDEBF7", text=("Start",))),
        _anchor((6, 8), (8, 12), _shape(4, "Rounded Rectangle 3", "roundRect", "E2EFDA", text=("Finish",))),
        _anchor((3, 10), (6, 10), (
            '<xdr:cxnSp macro=""><xdr:nvCxnSpPr><xdr:cNvPr id="5" name="Straight Arrow Connector 4"/>'
            '<xdr:cNvCxnSpPr><a:stCxn id="3" idx="3"/><a:endCxn id="4" idx="1"/></xdr:cNvCxnSpPr>'
            '</xdr:nvCxnSpPr><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm>'
            '<a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom>'
            '<a:ln w="19050"><a:solidFill><a:srgbClr val="305496"/></a:solidFill><a:tailEnd type="triangle"/>'
            '</a:ln></xdr:spPr></xdr:cxnSp>')),
        _anchor((1, 14), (5, 19), (
            '<xdr:grpSp><xdr:nvGrpSpPr><xdr:cNvPr id="8" name="Group 7"/><xdr:cNvGrpSpPr/></xdr:nvGrpSpPr>'
            '<xdr:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2400000" cy="900000"/>'
            '<a:chOff x="0" y="0"/><a:chExt cx="2400000" cy="900000"/></a:xfrm></xdr:grpSpPr>'
            + _shape(6, "Oval 5", "ellipse", "FCE4D6", text=("A",), size=(900000, 900000))
            + _shape(7, "Oval 6", "ellipse", "FFF2CC", text=("B",), offset=(1500000, 0), size=(900000, 900000))
            + '</xdr:grpSp>')),
        _anchor((9, 1), (12, 4), _shape(10, "Link 9", "roundRect", "DDEBF7", text=("Project page",), link="rId1")),
        _anchor((9, 8), (12, 12), _shape(11, "Picture fill 10", "ellipse", None, picture="rId1")),
        _anchor((6, 14), (9, 18), _shape(12, "Rotated 11", "star5", "FFD966", text=("Star",))
                .replace('<a:xfrm>', '<a:xfrm rot="1200000">', 1)
                .replace('</a:ln></xdr:spPr>', '</a:ln><a:effectLst><a:outerShdw blurRad="50800" dist="38100" '
                         'dir="2700000" algn="tl" rotWithShape="0"><a:prstClr val="black"><a:alpha val="40000"/>'
                         '</a:prstClr></a:outerShdw></a:effectLst></xdr:spPr>', 1)),
        _anchor((12, 21), (16, 25), _shape(13, "Callout 12", "wedgeRectCallout", "FFFFFF", text=(
            "Above the chart,", "below the picture"))),
    ],
    "Notes": [
        _anchor((1, 1), (6, 5), _shape(2, "TextBox 1", "rect", "FFF2CC", text_box=True, text=(
            "A sheet with shapes only:", "the drawing is created on import."))),
    ],
}


# Relationships and z-order of some shapes (see xlsx2txt.shapes).
_SHAPE_EXTRAS = {
    ("Diagram", 0): {"layer": 0},  # below the chart and the picture
    ("Diagram", 9): {"layer": 1},  # above the chart, below the picture
    ("Diagram", 6): {"rels": {"rId1": {"type": "hyperlink", "target": "https://github.com/click0/xlsx2txt",
                                       "external": True}}},
    ("Diagram", 7): {"rels": {"rId1": {"type": "image", "file": "fill.png"}}},
}
SHAPES = {
    sheet: [{"xml": xml, **_SHAPE_EXTRAS.get((sheet, number), {})} for number, xml in enumerate(fragments)]
    for sheet, fragments in _SHAPES_XML.items()
}


def shapes() -> tuple[Workbook, dict]:
    """Text box, shapes with text, a connector arrow, a group, a shape behind a
    picture, a shape with a hyperlink, one filled with a picture, a rotated
    shape with a shadow and a callout layered between a chart and a picture."""
    wb = _new_workbook("Shapes")
    ws = wb.active
    ws.title = "Diagram"
    ws["A1"] = "Process"
    ws.add_image(Image(io.BytesIO(_png((48, 84, 150), (48, 48)))), "A3")
    for row, (step, days) in enumerate([("Start", 2), ("Build", 5), ("Test", 3), ("Finish", 1)], start=21):
        ws.cell(row, 1, step)
        ws.cell(row, 2, days)
    chart = BarChart()
    chart.title = "Days per step"
    chart.add_data(Reference(ws, min_col=2, min_row=21, max_row=24))
    chart.set_categories(Reference(ws, min_col=1, min_row=21, max_row=24))
    chart.legend = None
    ws.add_chart(chart, "K20")
    wb.create_sheet("Notes")["A1"] = "See the note"
    return wb, {"shapes": SHAPES, "media": {"fill.png": _png((237, 125, 49), (16, 16))}}


# ---------------------------------------------------------------------------
# Worksheet extensions and threaded comments, as Excel writes them.
# ---------------------------------------------------------------------------

X14 = 'xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"'
XM = 'xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"'
DATA_BAR_ID = "{9F0C2E1A-5B7D-4C3E-8A21-6D4F0B9E7C15}"
SPARKLINE_COLORS = "".join(
    f'<x14:{name} rgb="{rgb}"/>' for name, rgb in [
        ("colorSeries", "FF376092"), ("colorNegative", "FFD00000"), ("colorAxis", "FF000000"),
        ("colorMarkers", "FFD00000"), ("colorFirst", "FFD00000"), ("colorLast", "FFD00000"),
        ("colorHigh", "FFD00000"), ("colorLow", "FFD00000"),
    ]
)
EXTENSIONS = [
    {"type": "conditional formatting", "xml": (
        f'<ext uri="{{78C0D931-6437-407d-A8EE-F0AAD7539E65}}" {X14}><x14:conditionalFormattings>'
        f'<x14:conditionalFormatting {XM}><x14:cfRule type="dataBar" id="{DATA_BAR_ID}">'
        '<x14:dataBar minLength="0" maxLength="100" gradient="0" negativeBarColorSameAsPositive="0" '
        'axisPosition="middle"><x14:cfvo type="autoMin"/><x14:cfvo type="autoMax"/>'
        '<x14:negativeFillColor rgb="FFFF0000"/><x14:axisColor rgb="FF000000"/></x14:dataBar></x14:cfRule>'
        '<xm:sqref>F2:F5</xm:sqref></x14:conditionalFormatting>'
        # An icon set that exists only as an extension: custom icons per level.
        f'<x14:conditionalFormatting {XM}><x14:cfRule type="iconSet" priority="2" '
        'id="{9F0C2E1A-5B7D-4C3E-8A21-6D4F0B9E7C16}"><x14:iconSet iconSet="3Stars" custom="1">'
        '<x14:cfvo type="percent"><xm:f>0</xm:f></x14:cfvo><x14:cfvo type="percent"><xm:f>33</xm:f></x14:cfvo>'
        '<x14:cfvo type="percent"><xm:f>67</xm:f></x14:cfvo><x14:cfIcon iconSet="3Flags" iconId="0"/>'
        '<x14:cfIcon iconSet="NoIcons" iconId="0"/><x14:cfIcon iconSet="3Stars" iconId="2"/></x14:iconSet>'
        '</x14:cfRule><xm:sqref>E2:E5</xm:sqref></x14:conditionalFormatting>'
        '</x14:conditionalFormattings></ext>')},
    {"type": "data validation", "xml": (
        f'<ext uri="{{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}}" {X14}><x14:dataValidations count="1" {XM}>'
        '<x14:dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1">'
        '<x14:formula1><xm:f>Lists!$A$1:$A$3</xm:f></x14:formula1><xm:sqref>H2:H5</xm:sqref>'
        '</x14:dataValidation></x14:dataValidations></ext>')},
    {"type": "sparklines", "xml": (
        f'<ext uri="{{05C60535-1F16-4fd2-B633-F4F36F0B64E0}}" {X14}><x14:sparklineGroups {XM}>'
        f'<x14:sparklineGroup lineWeight="1.5" displayEmptyCellsAs="gap" markers="1">{SPARKLINE_COLORS}<x14:sparklines>'
        + "".join(f"<x14:sparkline><xm:f>Sales!B{row}:E{row}</xm:f><xm:sqref>G{row}</xm:sqref></x14:sparkline>"
                  for row in range(2, 6))
        + '</x14:sparklines></x14:sparklineGroup>'
        f'<x14:sparklineGroup type="column" negative="1" high="1" displayEmptyCellsAs="gap">{SPARKLINE_COLORS}'
        '<x14:sparklines><x14:sparkline><xm:f>Sales!F2:F5</xm:f><xm:sqref>F7</xm:sqref></x14:sparkline>'
        '</x14:sparklines></x14:sparklineGroup>'
        f'<x14:sparklineGroup type="stacked" negative="1" displayEmptyCellsAs="gap">{SPARKLINE_COLORS}'
        '<x14:sparklines><x14:sparkline><xm:f>Sales!F2:F5</xm:f><xm:sqref>F8</xm:sqref></x14:sparkline>'
        '</x14:sparklines></x14:sparklineGroup></x14:sparklineGroups></ext>')},
]
PERSONS = [
    {"displayName": "Olena Koval", "id": "{4B1E7C2A-0D3F-4A5B-9C6D-7E8F9A0B1C2D}",
     "userId": "olena@example.com", "providerId": "None"},
    {"displayName": "Taras Melnyk", "id": "{5C2F8D3B-1E4A-4B6C-8D7E-9F0A1B2C3D4E}",
     "userId": "taras@example.com", "providerId": "None"},
]
THREAD_ID = "{6D3A9E4C-2F5B-4C7D-9E8F-0A1B2C3D4E5F}"
THREADS = [
    {"ref": "A3", "id": THREAD_ID, "personId": PERSONS[0]["id"], "dT": "2026-01-01T12:00:00.00",
     "text": "Is Q4 final?"},
    {"ref": "A3", "id": "{7E4B0F5D-3A6C-4D8E-8F9A-1B2C3D4E5F60}", "parentId": THREAD_ID,
     "personId": PERSONS[1]["id"], "dT": "2026-01-02T09:30:00.00", "text": "Yes, checked twice."},
    {"ref": "E4", "id": "{8F5C1A6E-4B7D-4E9F-9A0B-2C3D4E5F6071}", "personId": PERSONS[1]["id"],
     "dT": "2026-01-03T11:00:00.00", "done": "1", "text": "@Olena Koval East grew again, please check Q4",
     "xml": (f'<mentions xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">'
             f'<mention mentionpersonId="{PERSONS[0]["id"]}" '
             'mentionId="{9A6D2B7F-5C8E-4FA0-8B1C-3D4E5F607182}" startIndex="0" length="12"/></mentions>')},
]
LEGACY_NOTE = (
    "[Threaded comment]\n\nYour version of Excel allows you to read this threaded comment; however, any edits "
    "to it will get removed if the file is opened in a newer version of Excel. Learn more: "
    "https://go.microsoft.com/fwlink/?linkid=870924\n\nComment:\n    Is Q4 final?\nReply:\n    Yes, checked twice."
)


# Cell metadata as Excel writes it for dynamic array formulas.
DYNAMIC_ARRAY_METADATA = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
    'xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray">'
    '<metadataTypes count="1"><metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" '
    'pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" '
    'coerce="1" cellMeta="1"/></metadataTypes><futureMetadata name="XLDAPR" count="1"><bk><extLst>'
    '<ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}"><xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/>'
    '</ext></extLst></bk></futureMetadata><cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata>'
    '</metadata>'
)


def extensions() -> tuple[Workbook, dict]:
    """Sparklines, a data bar with extended options, a drop-down list from
    another sheet, a threaded comment with a reply, a dynamic array formula
    and a switched-off error check."""
    wb = _new_workbook("Extensions")
    ws = wb.active
    ws.title = "Sales"
    _header(ws, ["Region", "Q1", "Q2", "Q3", "Q4", "Change", "Trend", "Status"])
    for row, (region, *quarters) in enumerate([
        ("North", 120, 135, 128, 160), ("South", 90, 85, 70, 65),
        ("East", 60, 75, 95, 110), ("West", 140, 120, 150, 145),
    ], start=2):
        ws.append([region, *quarters, None, None, "Open"])
        ws[f"F{row}"] = f"=E{row}-B{row}"
    for column, width in zip("ABCDEFGH", (12, 8, 8, 8, 8, 10, 16, 12)):
        ws.column_dimensions[column].width = width
    ws.conditional_formatting.add("F2:F5", DataBarRule(start_type="min", end_type="max", color="FF638EC6"))
    ws["A3"].comment = Comment(LEGACY_NOTE, f"tc={THREAD_ID}")
    ws["E4"].comment = Comment("[Threaded comment]\n\nComment:\n    @Olena Koval East grew again, please check Q4",
                               "tc={8F5C1A6E-4B7D-4E9F-9A0B-2C3D4E5F6071}")
    ws["A7"], ws["A8"] = "Change (columns)", "Change (win/loss)"
    ws["J1"], ws["K1"] = "Sorted", "Code"
    ws["J2"] = ArrayFormula("J2:J5", "=_xlfn._xlws.SORT(A2:A5)")
    for row, region in enumerate(sorted(["North", "South", "East", "West"]), start=2):
        if row > 2:
            ws[f"J{row}"] = region  # the spilled values Excel stores
    ws["K2"] = "0042"  # a number stored as text; its error check is switched off
    lists = wb.create_sheet("Lists")
    for value in ("Open", "In progress", "Done"):
        lists.append([value])
    return wb, {
        "sheet_extensions": {"Sales": EXTENSIONS},
        "cf_ids": {"Sales": {"F2:F5#1": DATA_BAR_ID}},
        "sheet_threads": {"Sales": THREADS},
        "persons": PERSONS,
        "cell_metadata": {"Sales": {"J2": 1}},
        "metadata": DYNAMIC_ARRAY_METADATA,
        "sheet_elements": {"Sales": ['<ignoredErrors><ignoredError sqref="K2" numberStoredAsText="1"/></ignoredErrors>']},
    }


def complex_report() -> Workbook:
    """A small sales report: merged title, bold headers, formulas with a total,
    column widths and a taller title row (the file the CLI tests use)."""
    wb = _new_workbook("Complex")
    ws = wb.active
    ws.title = "Data"
    ws.merge_cells("A1:D1")
    ws["A1"] = "Sales Report Q1 2025"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A1"].alignment = Alignment(horizontal="center")
    for column, title in zip("ABCD", ("Product", "Quantity", "Price", "Total")):
        ws[f"{column}2"] = title
        ws[f"{column}2"].font = Font(bold=True)
    for row, (product, quantity, price) in enumerate([("Widget", 100, 9.99), ("Gadget", 50, 19.99),
                                                      ("Gizmo", 75, 14.99)], start=3):
        ws[f"A{row}"], ws[f"B{row}"], ws[f"C{row}"] = product, quantity, price
        ws[f"D{row}"] = f"=B{row}*C{row}"
    ws["A6"] = "Total"
    ws["A6"].font = Font(bold=True)
    ws["D6"] = "=SUM(D3:D5)"
    ws["D6"].font = Font(bold=True)
    for column, width in zip("ABCD", (15, 10, 10, 12)):
        ws.column_dimensions[column].width = width
    ws.row_dimensions[1].height = 25
    return wb


EXAMPLES = {
    "01-basic-data": basic_data,
    "02-styles": styles,
    "03-formulas": formulas,
    "04-images-charts": images_and_charts,
    "05-rich-text": rich_text,
    "06-workbook-features": workbook_features,
    "07-pivot-table": pivot_table,
    "08-shapes": shapes,
    "09-extensions": extensions,
    "10-complex": complex_report,
}


# ---------------------------------------------------------------------------
# Saving
# ---------------------------------------------------------------------------

_MODIFIED = re.compile(rb"<dcterms:modified([^>]*)>[^<]*</dcterms:modified>")


def save_deterministic(wb: Workbook, path: Path, patches: dict | None = None) -> None:
    """Save reproducibly: openpyxl stamps the document and zip entries with 'now'."""
    from xlsx2txt.package import patch_package

    wb.save(path)
    if patches:
        patch_package(path, **patches)
    buffer = io.BytesIO(path.read_bytes())
    stamp = FIXED_DATE.strftime("%Y-%m-%dT%H:%M:%SZ").encode()
    source = zipfile.ZipFile(io.BytesIO(buffer.getvalue()))
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as target:
        for item in source.infolist():
            data = source.read(item.filename)
            if item.filename == "docProps/core.xml":
                data = _MODIFIED.sub(lambda m: b"<dcterms:modified" + m.group(1) + b">" + stamp
                                     + b"</dcterms:modified>", data)
            info = zipfile.ZipInfo(item.filename, date_time=(1980, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            target.writestr(info, data)


def generate(target_dir: Path = EXAMPLES_DIR, export: bool = True) -> list:
    """Write every example workbook (and its export) into ``target_dir``."""
    from xlsx2txt import export_xlsx

    written = []
    for name, build in EXAMPLES.items():
        path = target_dir / f"{name}.xlsx"
        built = build()
        # Builders return the workbook, or the workbook and what to add to the
        # saved file that openpyxl cannot write (see xlsx2txt.package.patch_package).
        wb, patches = built if isinstance(built, tuple) else (built, None)
        save_deterministic(wb, path, patches)
        written.append(path)
        if export:
            export_xlsx(path, target_dir / name, force=True)
    return written


if __name__ == "__main__":
    sys.path.insert(0, str(EXAMPLES_DIR.parent))
    for written_path in generate():
        print(f"{written_path.relative_to(EXAMPLES_DIR.parent)} -> {written_path.stem}/")
