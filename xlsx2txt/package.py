"""Direct work with the xlsx package (zip) for parts openpyxl does not keep.

- Printer settings (``xl/printerSettings/*.bin``) are read from the source
  archive and written back into the saved file.
- Shapes are read from the drawings and written back into them
  (see :mod:`xlsx2txt.shapes`).
- A worksheet's background picture and the VML drawing with the pictures of
  its header/footer are read and written back.
- Worksheet extensions (sparklines, extended conditional formatting and data
  validation, see :mod:`xlsx2txt.extensions`) and threaded comments (see
  :mod:`xlsx2txt.threads`) are read from the parts and written back.
- Parts and sheet features that openpyxl drops are detected so that export
  can warn about them instead of losing them silently.
"""

import hashlib
import html
import posixpath
import re
import zipfile
from collections import Counter
from pathlib import Path

from openpyxl.packaging.relationship import get_dependents, get_rels_path
from openpyxl.reader.workbook import WorkbookParser

from xlsx2txt import elements, extensions, threads
from xlsx2txt.drawings import media_name
from xlsx2txt.xmlfrag import insert_child
from xlsx2txt.shapes import (
    EMPTY_DRAWING,
    add_to_drawing,
    classify,
    drawing_children,
    full_type,
    rewrite_rel_ids,
    shapes_from_drawing,
)

REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PRINTER_SETTINGS_REL = f"{REL_NS}/printerSettings"
PRINTER_SETTINGS_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings"
DRAWING_REL = f"{REL_NS}/drawing"
DRAWING_TYPE = "application/vnd.openxmlformats-officedocument.drawing+xml"
IMAGE_REL = f"{REL_NS}/image"
VML_REL = f"{REL_NS}/vmlDrawing"
VML_TYPE = "application/vnd.openxmlformats-officedocument.vmlDrawing"
_LEGACY_HF = re.compile(r"<(?:\w+:)?legacyDrawingHF\b[^>]*?\br:id=\"([^\"]+)\"")
# Pictures in VML: <v:imagedata o:relid="rId1"/> (or r:id).
_VML_REL = re.compile(r"\b(o:relid|r:id)=\"([^\"]*)\"")
_PICTURE = re.compile(r"<(?:\w+:)?picture\b[^>]*?\br:id=\"([^\"]+)\"")


def _sheet_paths(archive: zipfile.ZipFile) -> dict[str, str]:
    """Map sheet (and chart sheet) names to their part names."""
    parser = WorkbookParser(archive, "xl/workbook.xml")
    parser.parse()
    return {sheet.name: rel.target for sheet, rel in parser.find_sheets()}


def _sheet_drawings(archive: zipfile.ZipFile, names: set[str]) -> dict[str, list[str]]:
    """Map sheet (and chart sheet) names to the drawing parts they use."""
    result: dict[str, list[str]] = {}
    for sheet_name, sheet_path in _sheet_paths(archive).items():
        rels_path = get_rels_path(sheet_path)
        if sheet_path not in names or rels_path not in names:
            continue
        drawings = [rel.target for rel in get_dependents(archive, rels_path).find(DRAWING_REL)
                    if rel.target in names]
        if drawings:
            result[sheet_name] = drawings
    return result


def _background(archive: zipfile.ZipFile, names: set[str], sheet_path: str,
                sheet_xml: str) -> tuple[bytes, str] | None:
    """The background picture of a worksheet: ``(content, extension)``."""
    match = _PICTURE.search(sheet_xml)
    rels_path = get_rels_path(sheet_path)
    if not match or rels_path not in names:
        return None
    for rel in get_dependents(archive, rels_path).find(IMAGE_REL):
        if rel.Id == match.group(1) and rel.target in names and rel.TargetMode != "External":
            return archive.read(rel.target), posixpath.splitext(rel.target)[1].lstrip(".") or "png"
    return None


def _header_footer_pictures(archive: zipfile.ZipFile, names: set[str], sheet_path: str,
                            sheet_xml: str) -> tuple[dict, dict[str, bytes]] | None:
    """Pictures in the header/footer of a worksheet: ``({"xml": VML, "rels":
    {id: {"type": "image", "file": name}}}, media)``, or None if they cannot
    be kept."""
    match = _LEGACY_HF.search(sheet_xml)
    rels_path = get_rels_path(sheet_path)
    if not match or rels_path not in names:
        return None
    part = next((rel.target for rel in get_dependents(archive, rels_path).find(VML_REL)
                 if rel.Id == match.group(1)), None)
    if part not in names:
        return None
    try:
        xml = archive.read(part).decode("utf-8")
    except UnicodeDecodeError:  # kept as text; other encodings are reported instead
        return None
    part_rels_path = get_rels_path(part)
    part_rels = ({rel.Id: rel for rel in get_dependents(archive, part_rels_path)}
                 if part_rels_path in names else {})
    rels: dict[str, dict] = {}
    media: dict[str, bytes] = {}
    for _, rel_id in _VML_REL.findall(xml):
        rel = part_rels.get(rel_id)
        if rel is None or rel.Type != IMAGE_REL or rel.TargetMode == "External" or rel.target not in names:
            return None
        content = archive.read(rel.target)
        name = media_name(content, posixpath.splitext(rel.target)[1].lstrip(".") or "png")
        media[name] = content
        rels[rel_id] = {"type": "image", "file": name}
    return {"xml": xml, "rels": rels}, media


# ---------------------------------------------------------------------------
# Printer settings
# ---------------------------------------------------------------------------

def printer_settings_name(content: bytes) -> str:
    """Content-addressed file name; sheets often share identical settings."""
    return f"printer_{hashlib.sha256(content).hexdigest()[:16]}.bin"


def extract_printer_settings(path) -> dict[str, bytes]:
    """Return ``{sheet name: printer settings bytes}`` for worksheets that have them."""
    result: dict[str, bytes] = {}
    with zipfile.ZipFile(path) as archive:
        names = set(archive.namelist())
        for sheet_name, sheet_path in _sheet_paths(archive).items():
            rels_path = get_rels_path(sheet_path)
            if rels_path not in names or not sheet_path.startswith("xl/worksheets/"):
                continue
            for rel in get_dependents(archive, rels_path).find(PRINTER_SETTINGS_REL):
                if rel.target in names:
                    result[sheet_name] = archive.read(rel.target)
                    break
    return result


_RELATIONSHIPS_EMPTY = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'
)


def _add_relationship(rels_xml: str, rel_type: str, target: str, external: bool = False) -> tuple[str, str]:
    used = set(re.findall(r'\bId="([^"]+)"', rels_xml))
    number = 1
    while f"rId{number}" in used:
        number += 1
    rel_id = f"rId{number}"
    mode = ' TargetMode="External"' if external else ""
    element = f'<Relationship Id="{rel_id}" Type="{rel_type}" Target="{html.escape(target)}"{mode}/>'
    return rels_xml.replace("</Relationships>", element + "</Relationships>", 1), rel_id


def _declare_r(sheet_xml: str) -> str:
    root = re.search(r"<worksheet\b[^>]*>", sheet_xml)
    if root and "xmlns:r=" not in root.group(0):
        new_root = root.group(0)[:-1] + f' xmlns:r="{REL_NS}">'
        sheet_xml = sheet_xml.replace(root.group(0), new_root, 1)
    return sheet_xml


def _link_page_setup(sheet_xml: str, rel_id: str) -> str:
    """Point <pageSetup r:id=...> at the printer settings relationship."""
    sheet_xml = _declare_r(sheet_xml)
    setup = re.search(r"<pageSetup\b[^>]*?(/?)>", sheet_xml)
    if setup:
        tag = setup.group(0)
        tag = re.sub(r'\s+r:id="[^"]*"', "", tag)
        tag = re.sub(r"\s*(/?)>$", rf' r:id="{rel_id}"\1>', tag)
        return sheet_xml[:setup.start()] + tag + sheet_xml[setup.end():]
    margins = re.search(r"<pageMargins\b[^>]*/>", sheet_xml)
    if not margins:
        raise ValueError("worksheet without <pageMargins>: cannot attach printer settings")
    element = f'<pageSetup r:id="{rel_id}"/>'
    return sheet_xml[:margins.end()] + element + sheet_xml[margins.end():]


# Elements that follow <drawing> in a worksheet (ECMA-376, CT_Worksheet).
_AFTER_DRAWING = {"legacyDrawing", "legacyDrawingHF", "drawingHF", "picture", "oleObjects", "controls",
                  "webPublishItems", "tableParts", "extLst"}


def _link_drawing(sheet_xml: str, rel_id: str) -> str:
    return insert_child(_declare_r(sheet_xml), f'<drawing r:id="{rel_id}"/>', _AFTER_DRAWING)


def _free_part(parts: dict[str, bytes], pattern: str) -> str:
    number = 1
    while pattern.format(number) in parts:
        number += 1
    return pattern.format(number)


def _add_override(parts: dict[str, bytes], part: str, content_type: str) -> None:
    types = parts["[Content_Types].xml"].decode("utf-8")
    override = f'<Override PartName="/{part}" ContentType="{content_type}"/>'
    parts["[Content_Types].xml"] = types.replace("</Types>", override + "</Types>", 1).encode("utf-8")


def _add_sheet_relationship(parts: dict[str, bytes], sheet_path: str, rel_type: str, part: str) -> str:
    rels_path = get_rels_path(sheet_path)
    rels_xml = parts[rels_path].decode("utf-8") if rels_path in parts else _RELATIONSHIPS_EMPTY
    target = posixpath.relpath(part, posixpath.dirname(sheet_path))
    rels_xml, rel_id = _add_relationship(rels_xml, rel_type, target)
    parts[rels_path] = rels_xml.encode("utf-8")
    return rel_id


def _add_printer_settings(parts, sheet_paths, settings: dict[str, bytes]) -> None:
    for sheet_name, content in sorted(settings.items()):
        sheet_path = sheet_paths.get(sheet_name)
        if sheet_path is None:
            raise ValueError(f"no sheet named {sheet_name!r} for printer settings")
        part = _free_part(parts, "xl/printerSettings/printerSettings{}.bin")
        parts[part] = content
        rel_id = _add_sheet_relationship(parts, sheet_path, PRINTER_SETTINGS_REL, part)
        parts[sheet_path] = _link_page_setup(parts[sheet_path].decode("utf-8"), rel_id).encode("utf-8")
        _add_override(parts, part, PRINTER_SETTINGS_TYPE)


_IMAGE_TYPES = {
    "png": "image/png", "jpg": "image/jpeg", "jpeg": "image/jpeg", "gif": "image/gif", "bmp": "image/bmp",
    "tif": "image/tiff", "tiff": "image/tiff", "emf": "image/x-emf", "wmf": "image/x-wmf",
    "svg": "image/svg+xml", "wdp": "image/vnd.ms-photo", "vml": VML_TYPE,
}


def _ensure_default(parts: dict[str, bytes], extension: str) -> None:
    types = parts["[Content_Types].xml"].decode("utf-8")
    if re.search(rf'<Default\b[^>]*Extension="{re.escape(extension)}"', types, re.I):
        return
    content_type = _IMAGE_TYPES.get(extension.lower(), "application/octet-stream")
    default = f'<Default Extension="{extension}" ContentType="{content_type}"/>'
    types = re.sub(r"(<Types\b[^>]*>)", lambda m: m.group(1) + default, types, count=1)
    parts["[Content_Types].xml"] = types.encode("utf-8")


def _add_shapes(parts, sheet_paths, drawings, shapes: dict[str, list[dict]],
                media: dict[str, bytes]) -> None:
    media_parts: dict[str, str] = {}
    for sheet_name, sheet_shapes in sorted(shapes.items()):
        if not sheet_shapes:
            continue
        sheet_path = sheet_paths.get(sheet_name)
        if sheet_path is None:
            raise ValueError(f"no sheet named {sheet_name!r} for shapes")
        if sheet_name in drawings:
            part = drawings[sheet_name][0]
        else:
            part = _free_part(parts, "xl/drawings/drawing{}.xml")
            parts[part] = EMPTY_DRAWING.encode("utf-8")
            rel_id = _add_sheet_relationship(parts, sheet_path, DRAWING_REL, part)
            parts[sheet_path] = _link_drawing(parts[sheet_path].decode("utf-8"), rel_id).encode("utf-8")
            _add_override(parts, part, DRAWING_TYPE)

        rels_path = get_rels_path(part)
        rels_xml = parts[rels_path].decode("utf-8") if rels_path in parts else _RELATIONSHIPS_EMPTY
        placed = []
        for shape in sheet_shapes:
            mapping = {}
            for old_id, rel in (shape.get("rels") or {}).items():
                if rel.get("external"):
                    target, external = rel["target"], True
                else:
                    name = rel["file"]
                    if name not in media:
                        raise ValueError(f"Sheet '{sheet_name}': missing image file media/{name} of a shape")
                    if name not in media_parts:
                        extension = posixpath.splitext(name)[1].lstrip(".") or "png"
                        media_part = _free_part(parts, "xl/media/image{}." + extension)
                        parts[media_part] = media[name]
                        _ensure_default(parts, extension)
                        media_parts[name] = media_part
                    target = posixpath.relpath(media_parts[name], posixpath.dirname(part))
                    external = False
                rels_xml, mapping[old_id] = _add_relationship(rels_xml, full_type(rel["type"]), target, external)
            placed.append((rewrite_rel_ids(shape["xml"], mapping), shape.get("layer")))
        if rels_xml != _RELATIONSHIPS_EMPTY:
            parts[rels_path] = rels_xml.encode("utf-8")
        parts[part] = add_to_drawing(parts[part].decode("utf-8"), placed).encode("utf-8")


def _add_extensions(parts, sheet_paths, sheet_extensions: dict[str, list[dict]],
                    cf_ids: dict[str, dict[str, str]]) -> None:
    for sheet_name in sorted(set(sheet_extensions) | set(cf_ids)):
        exts = sheet_extensions.get(sheet_name) or []
        ids = cf_ids.get(sheet_name) or {}
        if not exts and not ids:
            continue
        sheet_path = sheet_paths.get(sheet_name)
        if sheet_path is None:
            raise ValueError(f"no sheet named {sheet_name!r} for extensions")
        xml = parts[sheet_path].decode("utf-8")
        parts[sheet_path] = extensions.write_sheet(xml, exts, ids).encode("utf-8")


def _add_sheet_xml(parts, sheet_paths, sheet_elements: dict[str, list[str]],
                   cell_metadata: dict[str, dict[str, int]], metadata: str | None) -> None:
    for sheet_name in sorted(set(sheet_elements) | set(cell_metadata)):
        if not sheet_elements.get(sheet_name) and not cell_metadata.get(sheet_name):
            continue
        sheet_path = sheet_paths.get(sheet_name)
        if sheet_path is None:
            raise ValueError(f"no sheet named {sheet_name!r} for sheet elements or cell metadata")
        xml = parts[sheet_path].decode("utf-8")
        if cell_metadata.get(sheet_name):
            xml = extensions.write_cell_metadata(xml, cell_metadata[sheet_name])
        if sheet_elements.get(sheet_name):
            xml = elements.write(xml, sheet_elements[sheet_name], elements.SHEET_ORDER)
        parts[sheet_path] = xml.encode("utf-8")
    if metadata:
        part = "xl/metadata.xml"
        parts[part] = metadata.encode("utf-8")
        _add_sheet_relationship(parts, "xl/workbook.xml", extensions.METADATA_REL, part)
        _add_override(parts, part, extensions.METADATA_TYPE)


def _add_backgrounds(parts, sheet_paths, backgrounds: dict[str, str], media: dict[str, bytes]) -> None:
    for sheet_name, name in sorted(backgrounds.items()):
        if not name:
            continue
        sheet_path = sheet_paths.get(sheet_name)
        if sheet_path is None:
            raise ValueError(f"no sheet named {sheet_name!r} for a background picture")
        if name not in media:
            raise ValueError(f"Sheet '{sheet_name}': missing background picture media/{name}")
        extension = posixpath.splitext(name)[1].lstrip(".") or "png"
        part = _free_part(parts, "xl/media/image{}." + extension)
        parts[part] = media[name]
        _ensure_default(parts, extension)
        rel_id = _add_sheet_relationship(parts, sheet_path, IMAGE_REL, part)
        after = set(elements.SHEET_ORDER[elements.SHEET_ORDER.index("picture") + 1:])
        xml = insert_child(_declare_r(parts[sheet_path].decode("utf-8")), f'<picture r:id="{rel_id}"/>', after)
        parts[sheet_path] = xml.encode("utf-8")


def _add_header_footer_pictures(parts, sheet_paths, pictures: dict[str, dict], media: dict[str, bytes]) -> None:
    for sheet_name, data in sorted(pictures.items()):
        if not data:
            continue
        sheet_path = sheet_paths.get(sheet_name)
        if sheet_path is None:
            raise ValueError(f"no sheet named {sheet_name!r} for header/footer pictures")
        part = _free_part(parts, "xl/drawings/vmlDrawingHF{}.vml")
        rels_xml = _RELATIONSHIPS_EMPTY
        mapping = {}
        for old_id, rel in sorted((data.get("rels") or {}).items()):
            name = rel["file"]
            if name not in media:
                raise ValueError(f"Sheet '{sheet_name}': missing header/footer picture media/{name}")
            extension = posixpath.splitext(name)[1].lstrip(".") or "png"
            media_part = _free_part(parts, "xl/media/image{}." + extension)
            parts[media_part] = media[name]
            _ensure_default(parts, extension)
            target = posixpath.relpath(media_part, posixpath.dirname(part))
            rels_xml, mapping[old_id] = _add_relationship(rels_xml, full_type(rel["type"]), target)
        xml = _VML_REL.sub(lambda m: f'{m.group(1)}="{mapping.get(m.group(2), m.group(2))}"', data["xml"])
        parts[part] = xml.encode("utf-8")
        if mapping:
            parts[get_rels_path(part)] = rels_xml.encode("utf-8")
        _ensure_default(parts, "vml")
        rel_id = _add_sheet_relationship(parts, sheet_path, VML_REL, part)
        after = set(elements.SHEET_ORDER[elements.SHEET_ORDER.index("legacyDrawingHF") + 1:])
        xml = insert_child(_declare_r(parts[sheet_path].decode("utf-8")), f'<legacyDrawingHF r:id="{rel_id}"/>',
                           after)
        parts[sheet_path] = xml.encode("utf-8")


def _zero_width_columns(parts, sheet_paths, columns: dict[str, list[int]]) -> None:
    """Hidden columns of width 0: openpyxl cannot write the width."""
    for sheet_name, indexes in sorted(columns.items()):
        if not indexes or sheet_name not in sheet_paths:
            continue
        wanted = {str(i) for i in indexes}
        xml = parts[sheet_paths[sheet_name]].decode("utf-8")

        def col(match: re.Match) -> str:
            tag = match.group(0)
            start = re.search(r'\bmin="(\d+)"', tag)
            if not start or start.group(1) not in wanted:
                return tag
            tag = re.sub(r'\s(width|customWidth)="[^"]*"', "", tag)
            end = len(tag) - (2 if tag.endswith("/>") else 1)
            return f'{tag[:end].rstrip()} width="0" customWidth="1"{tag[end:]}'
        parts[sheet_paths[sheet_name]] = re.sub(r"<(?:\w+:)?col\b[^>]*>", col, xml).encode("utf-8")


def _add_workbook_xml(parts, workbook_elements: list[str], workbook_extensions: list[dict]) -> None:
    if not workbook_elements and not workbook_extensions:
        return
    xml = parts["xl/workbook.xml"].decode("utf-8")
    xml = elements.write(xml, workbook_elements, elements.WORKBOOK_ORDER)
    xml = extensions.write_sheet(xml, workbook_extensions, {})
    parts["xl/workbook.xml"] = xml.encode("utf-8")


def _add_threads(parts, sheet_paths, sheet_threads: dict[str, list[dict]], persons: list[dict]) -> None:
    for sheet_name, records in sorted(sheet_threads.items()):
        if not records:
            continue
        sheet_path = sheet_paths.get(sheet_name)
        if sheet_path is None:
            raise ValueError(f"no sheet named {sheet_name!r} for threaded comments")
        part = _free_part(parts, "xl/threadedComments/threadedComment{}.xml")
        parts[part] = threads.write_threads(records).encode("utf-8")
        _add_sheet_relationship(parts, sheet_path, threads.THREAD_REL, part)
        _add_override(parts, part, threads.THREAD_TYPE)
    if persons:
        part = "xl/persons/person.xml"
        parts[part] = threads.write_persons(persons).encode("utf-8")
        _add_sheet_relationship(parts, "xl/workbook.xml", threads.PERSON_REL, part)
        _add_override(parts, part, threads.PERSON_TYPE)


def patch_package(path, printer_settings: dict[str, bytes] | None = None,
                  shapes: dict[str, list[dict]] | None = None, media: dict[str, bytes] | None = None,
                  sheet_extensions: dict[str, list[dict]] | None = None,
                  cf_ids: dict[str, dict[str, str]] | None = None,
                  sheet_threads: dict[str, list[dict]] | None = None,
                  persons: list[dict] | None = None,
                  sheet_elements: dict[str, list[str]] | None = None,
                  cell_metadata: dict[str, dict[str, int]] | None = None,
                  metadata: str | None = None,
                  workbook_elements: list[str] | None = None,
                  workbook_extensions: list[dict] | None = None,
                  zero_width_columns: dict[str, list[int]] | None = None,
                  backgrounds: dict[str, str] | None = None,
                  header_footer_pictures: dict[str, dict] | None = None) -> None:
    """Add what openpyxl does not write to a file it saved (in place).

    ``printer_settings``: ``{sheet name: bytes}``; ``shapes``: ``{sheet name:
    [shape, ...]}`` as exported (see :mod:`xlsx2txt.shapes`); ``media``: the
    pictures shapes refer to; ``sheet_extensions`` / ``cf_ids``: see
    :mod:`xlsx2txt.extensions`; ``sheet_threads`` / ``persons``: see
    :mod:`xlsx2txt.threads`; ``cell_metadata`` (``{sheet name: {cell: cm}}``)
    and ``metadata`` (the ``xl/metadata.xml`` part): see
    :mod:`xlsx2txt.extensions`; ``sheet_elements`` (``{sheet name: [XML]}``),
    ``workbook_elements`` (``[XML]``): see :mod:`xlsx2txt.elements`;
    ``workbook_extensions``: the workbook's ``<extLst>`` entries;
    ``zero_width_columns``: ``{sheet name: [column index]}`` of hidden columns
    with width 0; ``backgrounds``: ``{sheet name: media name}`` of background
    pictures; ``header_footer_pictures``: ``{sheet name: {"xml": VML,
    "rels": {...}}}``.
    """
    per_sheet = [shapes, sheet_extensions, cf_ids, sheet_threads, sheet_elements, cell_metadata,
                 zero_width_columns, backgrounds, header_footer_pictures]
    whole = [printer_settings, persons, metadata, workbook_elements, workbook_extensions]
    if not any(whole) and not any(any(m.values()) for m in per_sheet if m):
        return
    path = Path(path)
    with zipfile.ZipFile(path) as source:
        parts = {info.filename: source.read(info.filename) for info in source.infolist()}
        sheet_paths = _sheet_paths(source)
        drawings = _sheet_drawings(source, set(parts))

    _add_printer_settings(parts, sheet_paths, printer_settings or {})
    _add_shapes(parts, sheet_paths, drawings, shapes or {}, media or {})
    _add_extensions(parts, sheet_paths, sheet_extensions or {}, cf_ids or {})
    _add_threads(parts, sheet_paths, sheet_threads or {}, persons or [])
    _add_sheet_xml(parts, sheet_paths, sheet_elements or {}, cell_metadata or {}, metadata)
    _add_workbook_xml(parts, workbook_elements or [], workbook_extensions or [])
    _zero_width_columns(parts, sheet_paths, zero_width_columns or {})
    _add_backgrounds(parts, sheet_paths, backgrounds or {}, media or {})
    _add_header_footer_pictures(parts, sheet_paths, header_footer_pictures or {}, media or {})

    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as target_zip:
        # [Content_Types].xml first, as Office writes it.
        for name in ["[Content_Types].xml"] + [n for n in parts if n != "[Content_Types].xml"]:
            target_zip.writestr(name, parts[name])


def inject_printer_settings(path, settings: dict[str, bytes]) -> None:
    """Add printer settings to a file saved by openpyxl (in place)."""
    patch_package(path, printer_settings=settings)


def _drawing_shapes(archive: zipfile.ZipFile, names: set[str], drawing: str):
    rels_path = get_rels_path(drawing)
    rels = {rel.Id: rel for rel in get_dependents(archive, rels_path)} if rels_path in names else {}
    return shapes_from_drawing(archive.read(drawing).decode("utf-8", errors="replace"), rels, archive.read)


def extract_shapes(path) -> tuple[dict[str, list[dict]], dict[str, bytes]]:
    """Return ``({worksheet name: [shape, ...]}, {media name: bytes})``."""
    result: dict[str, list[dict]] = {}
    media: dict[str, bytes] = {}
    with zipfile.ZipFile(path) as archive:
        names = set(archive.namelist())
        sheet_paths = _sheet_paths(archive)
        for sheet_name, drawings in _sheet_drawings(archive, names).items():
            if not sheet_paths[sheet_name].startswith("xl/worksheets/"):
                continue
            for drawing in drawings:
                shapes, _, shape_media = _drawing_shapes(archive, names, drawing)
                result.setdefault(sheet_name, []).extend(shapes)
                media.update(shape_media)
    return result, media


def extract_sheet_extras(path) -> tuple[dict[str, dict], list[dict]]:
    """Extensions, conditional formatting links and threaded comments.

    Returns ``({worksheet name: {"extensions": [...], "cfIds": {...},
    "threadedComments": [...], "xmlElements": [...], "cellMetadata": {...},
    "background": (bytes, extension) | None, "headerFooterPictures": ({...},
    media) | None}},
    persons)``.
    """
    result: dict[str, dict] = {}
    persons: list[dict] = []
    with zipfile.ZipFile(path) as archive:
        names = set(archive.namelist())
        for sheet_name, sheet_path in _sheet_paths(archive).items():
            if sheet_path not in names or not sheet_path.startswith("xl/worksheets/"):
                continue
            sheet_xml = archive.read(sheet_path).decode("utf-8")
            exts, ids, _ = extensions.read_sheet(sheet_xml)
            records = []
            rels_path = get_rels_path(sheet_path)
            if rels_path in names:
                for rel in get_dependents(archive, rels_path).find(threads.THREAD_REL):
                    if rel.target in names:
                        records.extend(threads.read_threads(archive.read(rel.target).decode("utf-8")))
            result[sheet_name] = {
                "background": _background(archive, names, sheet_path, sheet_xml),
                "headerFooterPictures": _header_footer_pictures(archive, names, sheet_path, sheet_xml),
                "extensions": exts,
                "cfIds": ids,
                "threadedComments": records,
                "xmlElements": elements.read(sheet_xml, elements.SHEET_KEPT)[0],
                "cellMetadata": extensions.read_cell_metadata(sheet_xml),
            }
        workbook_rels = "xl/_rels/workbook.xml.rels"
        if workbook_rels in names:
            for rel in get_dependents(archive, workbook_rels).find(threads.PERSON_REL):
                if rel.target in names:
                    persons.extend(threads.read_persons(archive.read(rel.target).decode("utf-8")))
    return result, persons


def extract_workbook_extras(path) -> dict[str, list]:
    """Workbook elements and ``<extLst>`` entries openpyxl drops."""
    with zipfile.ZipFile(path) as archive:
        xml = archive.read("xl/workbook.xml").decode("utf-8")
    return {
        "xmlElements": elements.read(xml, elements.WORKBOOK_KEPT)[0],
        "extensions": extensions.read_sheet(xml)[0],
    }


def _keeps_metadata(xml: str) -> bool:
    """Cell metadata is kept unless it describes values (pictures in cells,
    rich data types), whose parts are not exported."""
    return "<valueMetadata" not in xml and ":valueMetadata" not in xml


def extract_metadata(path) -> str | None:
    """The ``xl/metadata.xml`` part (dynamic arrays), if it can be kept."""
    with zipfile.ZipFile(path) as archive:
        names = set(archive.namelist())
        workbook_rels = "xl/_rels/workbook.xml.rels"
        if workbook_rels not in names:
            return None
        for rel in get_dependents(archive, workbook_rels).find(extensions.METADATA_REL):
            if rel.target in names:
                xml = archive.read(rel.target).decode("utf-8")
                return xml if _keeps_metadata(xml) else None
    return None


# ---------------------------------------------------------------------------
# Parts and features that are not exported
# ---------------------------------------------------------------------------

# Parts xlsx2txt exports or that openpyxl regenerates from exported data.
_KNOWN_PARTS = re.compile(
    r"^(\[Content_Types\]\.xml|_rels/\.rels|docProps/(core|app|custom)\.xml"
    r"|xl/(workbook\.xml|_rels/workbook\.xml\.rels|styles\.xml|sharedStrings\.xml|calcChain\.xml"
    r"|theme/.*|worksheets/.*|chartsheets/.*|drawings/.*|comments\d*\.xml|comments/.*|media/.*|charts/.*"
    r"|tables/.*|pivotTables/.*|pivotCache/.*|externalLinks/.*|printerSettings/.*"
    r"|threadedComments/.*|persons/.*"
    r"|vbaProject\.bin|vbaProjectSignature.*\.bin|_rels/vbaProject\.bin\.rels))$"
)

# Recognised groups of parts that are lost, with a readable description.
_LOST_PARTS = [
    (re.compile(r"^xl/(slicers|slicerCaches)/"), "slicers"),
    (re.compile(r"^xl/timelines?|^xl/timelineCaches/"), "timelines"),
    (re.compile(r"^xl/(ctrlProps|activeX)/"), "form controls / ActiveX controls"),
    (re.compile(r"^xl/embeddings/"), "embedded objects (OLE)"),
    (re.compile(r"^xl/(connections\.xml|queryTables/)"), "data connections / queries"),
    (re.compile(r"^customXml/"), "custom XML parts (e.g. Power Query)"),
    (re.compile(r"^xl/model/"), "data model (Power Pivot)"),
    (re.compile(r"^xl/richData/"), "pictures in cells / rich data types"),
    (re.compile(r"^xl/metadata\.xml$"), "cell metadata of pictures in cells / rich data types"),
    (re.compile(r"^xl/webextensions/"), "Office add-ins"),
    (re.compile(r"^customUI"), "ribbon customisation"),
]

_SMART_ART = re.compile(r"drawingml/2006/diagram")


def unsupported_warnings(path, keep_vba: bool = False) -> list[str]:
    """Describe everything in the file that the export does not keep.

    ``keep_vba``: the VBA project and ribbon customisation are exported
    (macro-enabled workbooks).
    """
    warnings: list[str] = []
    with zipfile.ZipFile(path) as archive:
        names = archive.namelist()
        lost = Counter()
        unknown = []
        for name in names:
            # Folders and relationship files are covered by the parts they link.
            if name.endswith("/") or name.endswith(".rels") or _KNOWN_PARTS.match(name):
                continue
            if name == "xl/metadata.xml" and _keeps_metadata(archive.read(name).decode("utf-8", errors="replace")):
                continue
            match = next(((pattern, text) for pattern, text in _LOST_PARTS if pattern.match(name)), None)
            if match is None:
                unknown.append(name)
            elif match[1] and not (keep_vba and match[1] == "ribbon customisation"):
                lost[match[1]] += 1
        for description, count in sorted(lost.items()):
            warnings.append(f"{description}: {count} part(s) are not exported")
        if unknown:
            shown = ", ".join(unknown[:5]) + (" …" if len(unknown) > 5 else "")
            warnings.append(f"{len(unknown)} unknown part(s) are not exported: {shown}")

        if "xl/workbook.xml" in names:
            workbook_xml = archive.read("xl/workbook.xml").decode("utf-8", errors="replace")
            for name in elements.read(workbook_xml, elements.WORKBOOK_KEPT)[1]:
                warnings.append(f"Workbook: <{name}> refers to other parts and is not exported")

        drawings = _sheet_drawings(archive, set(names))
        for sheet_name, sheet_path in _sheet_paths(archive).items():
            if sheet_path not in names:
                continue
            xml = archive.read(sheet_path).decode("utf-8", errors="replace")
            is_worksheet = sheet_path.startswith("xl/worksheets/")
            for description in sorted(set(extensions.read_sheet(xml)[2])):
                warnings.append(f"Sheet '{sheet_name}': {description} are not exported")
            kept = set()
            if is_worksheet and _background(archive, set(names), sheet_path, xml):
                kept.add("picture")
            if is_worksheet and _header_footer_pictures(archive, set(names), sheet_path, xml):
                kept.add("legacyDrawingHF")
            for description in elements.linked(xml, kept):
                warnings.append(f"Sheet '{sheet_name}': {description} are not exported")
            for name in elements.read(xml, elements.SHEET_KEPT)[1]:
                warnings.append(f"Sheet '{sheet_name}': <{name}> refers to other parts and is not exported")
            shapes = 0
            smart_art = False
            for drawing_path in drawings.get(sheet_name, []):
                drawing = archive.read(drawing_path).decode("utf-8", errors="replace")
                if is_worksheet:
                    shapes += _drawing_shapes(archive, set(names), drawing_path)[1]
                else:
                    shapes += sum(1 for child in drawing_children(drawing)[1] if classify(child))
                smart_art = smart_art or bool(_SMART_ART.search(drawing))
            if shapes:
                what = "form control(s) or shape(s) with unsupported links" if is_worksheet else "shape(s)"
                warnings.append(f"Sheet '{sheet_name}': {shapes} {what} are not exported")
            if smart_art:
                warnings.append(f"Sheet '{sheet_name}': SmartArt diagrams are not exported")
    return warnings
