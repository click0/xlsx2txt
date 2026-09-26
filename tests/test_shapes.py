"""Shapes (text boxes, arrows, groups) are exported and restored."""

import re
import zipfile
from io import BytesIO

from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image
from PIL import Image as PILImage

from xlsx2txt import diff_models, export_model, export_xlsx, import_dir
from xlsx2txt.converter import roundtrip_diff
from xlsx2txt.shapes import drawing_children, shapes_from_drawing
from xlsx2txt.storage import write_model
from xlsx2txt.text import render_text

XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
A14 = "http://schemas.microsoft.com/office/drawing/2010/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _anchor(col, row, body):
    return (f"<xdr:twoCellAnchor><xdr:from><xdr:col>{col}</xdr:col><xdr:colOff>0</xdr:colOff>"
            f"<xdr:row>{row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>"
            f"<xdr:to><xdr:col>{col + 2}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{row + 2}</xdr:row>"
            f"<xdr:rowOff>0</xdr:rowOff></xdr:to>{body}<xdr:clientData/></xdr:twoCellAnchor>")


def _sp(shape_id, name, text, extra=""):
    return (f'<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="{shape_id}" name="{name}">{extra}'
            f'</xdr:cNvPr><xdr:cNvSpPr txBox="1"/></xdr:nvSpPr><xdr:spPr><a:prstGeom prst="rect"><a:avLst/>'
            f'</a:prstGeom></xdr:spPr><xdr:txBody><a:bodyPr/><a:lstStyle/>'
            f'<a:p><a:r><a:t>{text}</a:t></a:r></a:p></xdr:txBody></xdr:sp>')


# As Excel writes it: namespaces declared on the root only.
SHAPES = [
    _anchor(3, 1, _sp(1, "TextBox 1", "Hello &amp; welcome")),
    f'<mc:AlternateContent xmlns:mc="{MC}"><mc:Choice xmlns:a14="{A14}" Requires="a14">'
    + _anchor(3, 5, _sp(2, "Rectangle 2", "Choice")) + "</mc:Choice><mc:Fallback/></mc:AlternateContent>",
    _anchor(3, 9, '<xdr:cxnSp macro=""><xdr:nvCxnSpPr><xdr:cNvPr id="3" name="Arrow 3"/><xdr:cNvCxnSpPr>'
                  '<a:stCxn id="1" idx="2"/><a:endCxn id="2" idx="0"/></xdr:cNvCxnSpPr></xdr:nvCxnSpPr>'
                  '<xdr:spPr><a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom></xdr:spPr></xdr:cxnSp>'),
]
LINKED = _anchor(8, 1, _sp(4, "Link 4", "Click", '<a:hlinkClick r:id="rId9"/>'))


def _png():
    buffer = BytesIO()
    PILImage.new("RGB", (4, 4), "red").save(buffer, format="PNG")
    return buffer.getvalue()


def _excel_like(path):
    """A picture (id 1 after openpyxl) plus shapes whose ids clash with it."""
    wb = Workbook()
    wb.active.title = "Data"
    wb.active["A1"] = "value"
    wb.active.add_image(Image(BytesIO(_png())), "A3")
    wb.create_sheet("Empty")
    wb.save(path)
    with zipfile.ZipFile(path) as source:
        parts = {name: source.read(name) for name in source.namelist()}
    drawing = parts["xl/drawings/drawing1.xml"].decode()
    root, children = drawing_children(drawing)
    new_root = (f'<xdr:wsDr xmlns:xdr="{XDR}" xmlns:a="{A}" xmlns:r="{R}">')
    # Re-prefix openpyxl's picture anchor so the drawing looks like Excel's.
    picture = re.sub(r"<(/?)(?!a:)(\w+)", r"<\1xdr:\2", children[0])
    picture = re.sub(r"<(/?)xdr:(blip|stretch|fillRect|prstGeom|avLst|off|ext|xfrm|picLocks)\b", r"<\1a:\2",
                     picture)
    parts["xl/drawings/drawing1.xml"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + new_root + picture
        + "".join(SHAPES) + LINKED + "</xdr:wsDr>"
    ).encode()
    with zipfile.ZipFile(path, "w") as target:
        for name, content in parts.items():
            target.writestr(name, content)


def test_drawing_children_and_classification():
    xml = f'<xdr:wsDr xmlns:xdr="{XDR}" xmlns:a="{A}" xmlns:r="{R}">' + "".join(SHAPES) + LINKED + "</xdr:wsDr>"
    root, children = drawing_children(xml)
    assert root.startswith("<xdr:wsDr") and len(children) == 4
    shapes, linked = shapes_from_drawing(xml)
    assert linked == 1
    assert [s["name"] for s in shapes] == ["TextBox 1", "Rectangle 2", "Arrow 3"]
    assert shapes[0]["text"] == "Hello & welcome"
    assert "text" not in shapes[2]
    # Every fragment declares the namespaces it uses.
    assert shapes[0]["xml"].startswith(f'<xdr:twoCellAnchor xmlns:xdr="{XDR}" xmlns:a="{A}">')
    assert f'xmlns:mc="{MC}"' in shapes[1]["xml"] and f'xmlns:xdr="{XDR}"' in shapes[1]["xml"]


def test_shapes_roundtrip(tmp_path):
    source = tmp_path / "shapes.xlsx"
    _excel_like(source)
    out_dir = tmp_path / "out"
    model = export_xlsx(source, out_dir)
    sheet = model["sheets"][0]
    assert [s["name"] for s in sheet["shapes"]] == ["TextBox 1", "Rectangle 2", "Arrow 3"]
    assert len(sheet["images"]) == 1
    assert model["manifest"]["warnings"] == [
        "Sheet 'Data': 1 shape(s) with pictures, links or controls are not exported"
    ]
    assert roundtrip_diff(model) == []

    restored = tmp_path / "restored.xlsx"
    import_dir(out_dir, restored)
    with zipfile.ZipFile(restored) as archive:
        drawing = archive.read("xl/drawings/drawing1.xml").decode()
    ids = re.findall(r'cNvPr id="(\d+)"', drawing)
    assert len(ids) == len(set(ids)) == 4  # picture + 3 shapes, clash with the picture resolved
    # The connector still points at the renumbered shapes.
    start, end = re.search(r'stCxn id="(\d+)".*?endCxn id="(\d+)"', drawing).groups()
    names = dict(re.findall(r'cNvPr id="(\d+)" name="([^"]+)"', drawing))
    assert (names[start], names[end]) == ("TextBox 1", "Rectangle 2")
    load_workbook(restored)  # still readable
    again = export_model(restored)
    assert diff_models(model, again, ignore_cached=True) == []
    assert [s["text"] for s in again["sheets"][0]["shapes"][:2]] == ["Hello & welcome", "Choice"]


def test_shapes_on_sheet_without_drawing(tmp_path):
    source = tmp_path / "shapes.xlsx"
    _excel_like(source)
    out_dir = tmp_path / "out"
    model = export_xlsx(source, out_dir)
    # Move the shapes to the sheet that has no drawing yet.
    model["sheets"][1]["shapes"] = model["sheets"][0].pop("shapes")
    write_model(model, out_dir)
    restored = tmp_path / "restored.xlsx"
    import_dir(out_dir, restored)
    with zipfile.ZipFile(restored) as archive:
        sheet = archive.read("xl/worksheets/sheet2.xml").decode()
        types = archive.read("[Content_Types].xml").decode()
        assert re.search(r'<drawing r:id="rId\d+"/></worksheet>', sheet)
        assert types.count("drawing+xml") == 2
    again = export_model(restored)
    assert "shapes" not in again["sheets"][0]
    assert [s["name"] for s in again["sheets"][1]["shapes"]] == ["TextBox 1", "Rectangle 2", "Arrow 3"]


def test_shape_changes_in_diff_and_text(tmp_path):
    source = tmp_path / "shapes.xlsx"
    _excel_like(source)
    model = export_model(source)
    text = render_text(model)
    assert "shape 'TextBox 1' at D2: 'Hello & welcome'" in text
    assert "shape 'Arrow 3' at D10" in text

    edited = export_model(source)
    shape = edited["sheets"][0]["shapes"][0]
    shape["xml"] = shape["xml"].replace("Hello &amp; welcome", "Bye")
    shape["text"] = "Bye"
    del edited["sheets"][0]["shapes"][2]
    assert diff_models(model, edited) == [
        "[Data] shape 'TextBox 1': text changed",
        "[Data] shape 'Arrow 3': removed",
    ]
