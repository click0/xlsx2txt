"""Semantic comparison of two xlsx2txt models."""

from typing import Any

from xlsx2txt.shapes import normalized as normalized_shapes

# Properties rewritten by Excel/openpyxl on every save.
VOLATILE_PROPERTIES = {"modified"}

_PART_KEYS = {
    "fonts": "font",
    "fills": "fill",
    "borders": "border",
    "alignments": "alignment",
    "protections": "protection",
}


def resolve_style(styles: dict[str, Any], index: int | None) -> dict[str, Any]:
    """Expand a cell style index into a full style description."""
    cell_styles = styles.get("cellStyles") or []
    if not cell_styles:
        return {}
    style = cell_styles[index or 0]
    result = {key: styles[part][style[key]] for part, key in _PART_KEYS.items()}
    result["numFmt"] = style.get("numFmt", "General")
    return result


def _fmt(value: Any) -> str:
    text = repr(value)
    return text if len(text) <= 60 else text[:57] + "..."


def _diff_values(path: str, a: Any, b: Any, out: list[str]) -> None:
    if a == b:
        return
    if isinstance(a, dict) and isinstance(b, dict):
        for key in list(a) + [k for k in b if k not in a]:
            if key not in b:
                out.append(f"{path}.{key}: removed (was {_fmt(a[key])})")
            elif key not in a:
                out.append(f"{path}.{key}: added {_fmt(b[key])}")
            else:
                _diff_values(f"{path}.{key}", a[key], b[key], out)
        return
    out.append(f"{path}: {_fmt(a)} -> {_fmt(b)}")


def _cell_view(record: dict[str, Any] | None, styles: dict[str, Any], ignore_cached: bool) -> dict[str, Any]:
    record = dict(record or {})
    index = record.pop("s", 0)
    if ignore_cached and record.get("t") == "f":
        record.pop("v", None)
        record.pop("vt", None)
    record["style"] = resolve_style(styles, index)
    return record


def _dims_view(dims: dict[str, Any], styles: dict[str, Any]) -> dict[str, Any]:
    result = {}
    for kind in ("columns", "rows"):
        entries = {}
        for key, data in dims.get(kind, {}).items():
            data = dict(data)
            # openpyxl cannot write a zero width: a hidden 0-width column comes
            # back with the default width 13. Both look the same in Excel.
            if kind == "columns" and data.get("hidden") and data.get("width") in (0, 13):
                data.pop("width")
            if "s" in data:
                data["style"] = resolve_style(styles, data.pop("s"))
            entries[key] = data
        result[kind] = entries
    return result


def _diff_shapes(prefix: str, shapes_a: list, shapes_b: list, out: list[str]) -> None:
    xml_a = normalized_shapes(shapes_a)
    xml_b = normalized_shapes(shapes_b)
    names_a = [s.get("name") for s in shapes_a]
    names_b = [s.get("name") for s in shapes_b]
    # Pair shapes by name when names identify them, otherwise by position.
    if None not in names_a + names_b and len(set(names_a)) == len(names_a) and len(set(names_b)) == len(names_b):
        keys_a, keys_b = names_a, names_b
    else:
        keys_a, keys_b = list(range(len(shapes_a))), list(range(len(shapes_b)))
    index_b = {key: i for i, key in enumerate(keys_b)}
    for i, key in enumerate(keys_a):
        label = f"{prefix} shape {shapes_a[i].get('name', i + 1)!r}"
        if key not in index_b:
            out.append(f"{label}: removed")
            continue
        j = index_b[key]
        if xml_a[i] != xml_b[j]:
            same_text = shapes_a[i].get("text") == shapes_b[j].get("text")
            out.append(f"{label}: {'changed' if same_text else 'text changed'}")
    for j, key in enumerate(keys_b):
        if key not in keys_a:
            out.append(f"{prefix} shape {shapes_b[j].get('name', j + 1)!r}: added")


def _diff_threads(prefix: str, threads_a: list, threads_b: list, out: list[str]) -> None:
    by_id_a = {t.get("id"): t for t in threads_a}
    by_id_b = {t.get("id"): t for t in threads_b}
    for key in list(by_id_a) + [k for k in by_id_b if k not in by_id_a]:
        a, b = by_id_a.get(key), by_id_b.get(key)
        label = f"{prefix} comment at {(a or b).get('ref')} {key}"
        if b is None:
            out.append(f"{label}: removed")
        elif a is None:
            out.append(f"{label}: added {_fmt(b.get('text'))}")
        elif a.get("text") != b.get("text"):
            out.append(f"{label}: {_fmt(a.get('text'))} -> {_fmt(b.get('text'))}")
        elif a != b:
            _diff_values(label, {k: v for k, v in a.items() if k != "text"},
                         {k: v for k, v in b.items() if k != "text"}, out)


def _diff_extensions(prefix: str, exts_a: list, exts_b: list, out: list[str]) -> None:
    for i in range(max(len(exts_a), len(exts_b))):
        a = exts_a[i] if i < len(exts_a) else None
        b = exts_b[i] if i < len(exts_b) else None
        label = f"{prefix} extension {(a or b).get('type')!r}"
        if b is None:
            out.append(f"{label}: removed")
        elif a is None:
            out.append(f"{label}: added")
        elif a != b:
            out.append(f"{label}: changed")


def _diff_sheet(name: str, a: dict[str, Any], b: dict[str, Any],
                styles_a: dict[str, Any], styles_b: dict[str, Any],
                ignore_cached: bool, out: list[str]) -> None:
    prefix = f"[{name}]"
    skip = {"cells", "dimensions", "name", "shapes", "extensions", "threadedComments"}
    for key in list(a) + [k for k in b if k not in a]:
        if key in skip:
            continue
        _diff_values(f"{prefix} {key}", a.get(key), b.get(key), out)

    _diff_shapes(prefix, a.get("shapes", []), b.get("shapes", []), out)
    _diff_threads(prefix, a.get("threadedComments", []), b.get("threadedComments", []), out)
    _diff_extensions(prefix, a.get("extensions", []), b.get("extensions", []), out)

    _diff_values(
        f"{prefix} dimensions",
        _dims_view(a.get("dimensions", {}), styles_a),
        _dims_view(b.get("dimensions", {}), styles_b),
        out,
    )

    cells_a = a.get("cells", {})
    cells_b = b.get("cells", {})
    default_a = _cell_view(None, styles_a, ignore_cached)
    default_b = _cell_view(None, styles_b, ignore_cached)
    for coord in list(cells_a) + [c for c in cells_b if c not in cells_a]:
        view_a = _cell_view(cells_a.get(coord), styles_a, ignore_cached)
        view_b = _cell_view(cells_b.get(coord), styles_b, ignore_cached)
        if view_a == view_b:
            continue
        if coord not in cells_a and view_a == default_a:
            out.append(f"{prefix} {coord}: added {_fmt(cells_b[coord].get('f', cells_b[coord].get('v')))}")
        elif coord not in cells_b and view_b == default_b:
            out.append(f"{prefix} {coord}: removed (was {_fmt(cells_a[coord].get('f', cells_a[coord].get('v')))})")
        else:
            _diff_values(f"{prefix} {coord}", view_a, view_b, out)


def diff_models(a: dict[str, Any], b: dict[str, Any], ignore_cached: bool = False) -> list[str]:
    """Return a human readable list of differences between two models."""
    out: list[str] = []

    wb_a = dict(a["workbook"])
    wb_b = dict(b["workbook"])
    props_a = {k: v for k, v in wb_a.pop("properties", {}).items() if k not in VOLATILE_PROPERTIES}
    props_b = {k: v for k, v in wb_b.pop("properties", {}).items() if k not in VOLATILE_PROPERTIES}
    _diff_values("workbook.properties", props_a, props_b, out)
    _diff_values("workbook", wb_a, wb_b, out)

    sheets_a = {s["name"]: s for s in a["sheets"]}
    sheets_b = {s["name"]: s for s in b["sheets"]}
    for name in sheets_a:
        if name not in sheets_b:
            out.append(f"[{name}] sheet removed")
    for name in sheets_b:
        if name not in sheets_a:
            out.append(f"[{name}] sheet added")
    for name in sheets_a:
        if name in sheets_b:
            _diff_sheet(name, sheets_a[name], sheets_b[name], a["styles"], b["styles"], ignore_cached, out)

    if (a.get("theme") or None) != (b.get("theme") or None):
        out.append("theme: changed")

    media_a = a.get("media") or {}
    media_b = b.get("media") or {}
    for name in sorted(set(media_a) | set(media_b)):
        if media_a.get(name) != media_b.get(name):
            state = "added" if name not in media_a else "removed" if name not in media_b else "changed"
            out.append(f"media/{name}: {state}")

    printer_a = a.get("printerSettings") or {}
    printer_b = b.get("printerSettings") or {}
    for name in sorted(set(printer_a) | set(printer_b)):
        if printer_a.get(name) != printer_b.get(name):
            state = "added" if name not in printer_a else "removed" if name not in printer_b else "changed"
            out.append(f"printer/{name}: {state}")

    pivots_a = a.get("pivots") or {}
    pivots_b = b.get("pivots") or {}
    for name in sorted(set(pivots_a) | set(pivots_b)):
        if pivots_a.get(name) != pivots_b.get(name):
            state = "added" if name not in pivots_a else "removed" if name not in pivots_b else "changed"
            out.append(f"pivots/{name}: {state}")

    src_a = a.get("vbaSources") or {}
    src_b = b.get("vbaSources") or {}
    for name in sorted(set(src_a) | set(src_b)):
        if src_a.get(name) != src_b.get(name):
            state = "added" if name not in src_a else "removed" if name not in src_b else "changed"
            out.append(f"vba/modules/{name}: {state}")

    vba_a = a.get("vba") or {}
    vba_b = b.get("vba") or {}
    for name in sorted(set(vba_a) | set(vba_b)):
        if vba_a.get(name) != vba_b.get(name):
            state = "added" if name not in vba_a else "removed" if name not in vba_b else "changed"
            out.append(f"vba/{name}: {state}")

    return out


def validate_model(model: dict[str, Any]) -> list[str]:
    """Check internal consistency of a model (style references, names...)."""
    errors = []
    styles = model.get("styles", {})
    cell_styles = styles.get("cellStyles", [])
    for i, style in enumerate(cell_styles):
        for part, key in _PART_KEYS.items():
            ref = style.get(key)
            if not isinstance(ref, int) or not 0 <= ref < len(styles.get(part, [])):
                errors.append(f"cellStyles[{i}].{key}: invalid reference {ref!r}")

    names = [s.get("name") for s in model.get("sheets", [])]
    chartsheet_names = [c.get("name") for c in model.get("workbook", {}).get("chartsheets", [])]
    all_names = names + chartsheet_names
    if len({n.lower() for n in all_names if n}) != len(all_names):
        errors.append("duplicate or empty sheet names")
    listed = model.get("workbook", {}).get("sheets")
    if listed is not None and listed != names:
        errors.append("workbook.sheets does not match the sheet index")

    valid_types = {"s", "rs", "n", "b", "e", "d", "td", "f"}
    media = model.get("media") or {}
    pivots = model.get("pivots") or {}
    for sheet in model.get("sheets", []):
        printer_file = sheet.get("printerSettings")
        if printer_file and printer_file not in (model.get("printerSettings") or {}):
            errors.append(f"[{sheet.get('name')}] printer settings not found: data/printer/{printer_file}")
        for entry in sheet.get("pivotTables", []):
            names = [entry.get("table")]
            if entry.get("cache"):
                names.append(f"{entry['cache']}.xml")
            for name in names:
                if name not in pivots:
                    errors.append(f"[{sheet.get('name')}] pivot file not found: data/pivots/{name}")
        person_ids = {p.get("id") for p in model.get("workbook", {}).get("persons") or []}
        for record in sheet.get("threadedComments", []):
            if record.get("personId") not in person_ids:
                errors.append(f"[{sheet.get('name')}] comment {record.get('id')}: unknown person "
                              f"{record.get('personId')!r} (workbook.persons)")
        for number, shape in enumerate(sheet.get("shapes", []), start=1):
            if not isinstance(shape.get("xml"), str) or not shape["xml"].startswith("<"):
                errors.append(f"[{sheet.get('name')}] shape {number}: no XML")
            for rel in (shape.get("rels") or {}).values():
                if not rel.get("external") and rel.get("file") not in media:
                    errors.append(f"[{sheet.get('name')}] shape {number}: file not found: data/media/{rel.get('file')}")
        for image in sheet.get("images", []):
            if image.get("file") not in media:
                errors.append(f"[{sheet.get('name')}] image file not found: data/media/{image.get('file')}")
        for coord, record in sheet.get("cells", {}).items():
            where = f"[{sheet.get('name')}] {coord}"
            index = record.get("s", 0)
            if not isinstance(index, int) or not 0 <= index < max(len(cell_styles), 1):
                errors.append(f"{where}: invalid style index {index!r}")
            value_type = record.get("t")
            if value_type is not None and value_type not in valid_types:
                errors.append(f"{where}: unknown type {value_type!r}")
            if value_type == "f" and "f" not in record and record.get("fType") != "dataTable":
                errors.append(f"{where}: formula cell without formula")
    return errors
