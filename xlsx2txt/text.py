"""Plain-text rendering of a model, one line per cell.

Used by ``xlsx2txt cat`` so that ``git diff`` can show changes inside
Excel files directly (see "Git integration" in the README).
"""

from typing import Any

from xlsx2txt.compare import resolve_style
from xlsx2txt.drawings import chart_title
from xlsx2txt.pivots import describe as describe_pivot
from xlsx2txt.shapes import shape_cell


def _color(value: Any) -> str:
    if isinstance(value, str):
        return value
    if isinstance(value, dict):
        return ",".join(f"{k}={v}" for k, v in value.items())
    return str(value)


def describe_style(style: dict[str, Any], default: dict[str, Any]) -> str:
    """Short human readable description of the parts that differ from the default."""
    parts: list[str] = []
    font = style.get("font", {})
    if font != default.get("font"):
        bits = []
        if font.get("name") and font.get("name") != default.get("font", {}).get("name"):
            bits.append(font["name"])
        if font.get("size") and font.get("size") != default.get("font", {}).get("size"):
            bits.append(f"{font['size']:g}pt")
        for flag in ("bold", "italic", "strike"):
            if font.get(flag):
                bits.append(flag)
        if font.get("underline"):
            bits.append("underline")
        if "color" in font and font.get("color") != default.get("font", {}).get("color"):
            bits.append(f"color {_color(font['color'])}")
        if bits:
            parts.append("font " + " ".join(bits))
    fill = style.get("fill", {})
    if fill != default.get("fill"):
        if "gradient" in fill:
            parts.append("fill gradient")
        elif fill.get("patternType"):
            color = fill.get("fgColor")
            parts.append(f"fill {fill['patternType']}" + (f" {_color(color)}" if color else ""))
    border = style.get("border", {})
    if border != default.get("border"):
        sides = [f"{side}:{data.get('style')}" for side, data in border.items() if isinstance(data, dict)]
        if sides:
            parts.append("border " + " ".join(sides))
    alignment = style.get("alignment", {})
    if alignment != default.get("alignment"):
        parts.append("align " + " ".join(f"{k}={v}" for k, v in alignment.items()))
    if style.get("numFmt", "General") != default.get("numFmt", "General"):
        parts.append(f"format {style['numFmt']}")
    protection = style.get("protection", {})
    if protection != default.get("protection"):
        parts.append("protection " + " ".join(f"{k}={v}" for k, v in protection.items()))
    return "; ".join(parts)


def _value(record: dict[str, Any]) -> str:
    value_type = record.get("t")
    if value_type == "f":
        text = record.get("f") or f"<{record.get('fType', 'formula')}>"
        if record.get("fType") == "array":
            text = f"{{{text}}} over {record.get('fRef')}"
        if "v" in record:
            text += f"  -> {record['v']!r}"
        return text
    if value_type == "rs":
        return repr("".join(run if isinstance(run, str) else run.get("text", "") for run in record["v"])) + " (rich text)"
    if value_type == "e":
        return str(record.get("v"))
    if "v" in record:
        return repr(record["v"])
    return ""


def render_text(model: dict[str, Any], styles: bool = True) -> str:
    """Render a model as stable, diff-friendly text."""
    lines: list[str] = []
    style_table = model["styles"]
    default = resolve_style(style_table, 0)
    workbook = model["workbook"]

    names = workbook.get("definedNames", [])
    for name in names:
        lines.append(f"name {name['name']} = {name['value']}")
    for index, link in enumerate(workbook.get("externalLinks", []), 1):
        lines.append(f"external link [{index}] -> {link['target']}")
    if lines:
        lines.append("")

    for sheet in model["sheets"]:
        header = f"=== {sheet['name']}"
        if sheet.get("state"):
            header += f" [{sheet['state']}]"
        lines.append(header + " ===")
        for rng in sheet.get("mergedCells", []):
            lines.append(f"merged {rng}")
        for image in sheet.get("images", []):
            anchor = image["anchor"]
            where = anchor.get("from", {}).get("cell", anchor.get("type"))
            lines.append(f"image at {where}: {image['file']}")
        for chart in sheet.get("charts", []):
            where = chart["anchor"].get("from", {}).get("cell", chart["anchor"].get("type"))
            title = chart_title(chart)
            lines.append(f"chart at {where}" + (f": {title}" if title else ""))
        for shape in sheet.get("shapes", []):
            line = f"shape {shape.get('name', '')!r} at {shape_cell(shape) or '?'}"
            if shape.get("text"):
                line += f": {shape['text']!r}"
            lines.append(line)
        for entry in sheet.get("pivotTables", []):
            description = describe_pivot((model.get("pivots") or {}).get(entry["table"], ""))
            lines.append(f"pivot table {description or entry['table']}")
        for coord, record in sheet.get("cells", {}).items():
            value = _value(record)
            line = f"{coord}: {value}" if value else f"{coord}:"
            if styles and record.get("s"):
                description = describe_style(resolve_style(style_table, record["s"]), default)
                if description:
                    line += f"  [{description}]"
            if "link" in record:
                line += f"  <link {record['link'].get('target') or record['link'].get('location')}>"
            if "comment" in record:
                line += f"  # {record['comment'].get('author', '')}: {record['comment'].get('text', '')!r}"
            lines.append(line)
        lines.append("")

    for chartsheet in workbook.get("chartsheets", []):
        lines.append(f"=== {chartsheet['name']} (chart sheet) ===")
        for chart in chartsheet.get("charts", []):
            title = chart_title(chart)
            lines.append("chart" + (f": {title}" if title else ""))
        lines.append("")

    for name in sorted(model.get("vbaSources") or {}):
        lines.append(f"=== VBA {name} ===")
        lines.extend(model["vbaSources"][name].rstrip("\n").split("\n"))
        lines.append("")

    return "\n".join(lines).rstrip("\n") + "\n"
