"""Reading and writing the xlsx2txt directory layout.

Layout::

    manifest.json
    data/workbook.json
    data/sheets/_index.json
    data/sheets/<Sheet>.json
    data/styles/{fonts,fills,borders,alignments,protections,cellStyles}.json
    data/theme/theme1.xml          (optional)
    data/media/image_<hash>.<ext>  (optional, images)
    data/pivots/*.xml              (optional, pivot tables and their caches)
    data/printer/printer_<hash>.bin (optional, printer driver settings)
    data/metadata.xml              (optional, cell metadata: dynamic arrays)
    vba/xl/vbaProject.bin          (optional, .xlsm only)
    vba/modules/*.bas|cls|frm      (optional, VBA source code, read-only)
    _verify/checksums.json
"""

import hashlib
import json
import shutil
from pathlib import Path
from typing import Any

from xlsx2txt.exporter import FORMAT_NAME, FORMAT_VERSION
from xlsx2txt.names import safe_file_name
from xlsx2txt.styles import STYLE_PARTS

MANIFEST = "manifest.json"
WORKBOOK = "data/workbook.json"
SHEETS_DIR = "data/sheets"
SHEETS_INDEX = "data/sheets/_index.json"
STYLES_DIR = "data/styles"
THEME = "data/theme/theme1.xml"
MEDIA_DIR = "data/media"
PIVOTS_DIR = "data/pivots"
PRINTER_DIR = "data/printer"
METADATA = "data/metadata.xml"
VBA_DIR = "vba"
VBA_SOURCES_DIR = "vba/modules"
CHECKSUMS = "_verify/checksums.json"

MANAGED = [MANIFEST, "data", VBA_DIR, "_verify"]

_LINE_WIDTH = 100
# Objects with scalar values only (e.g. cells) stay on one line up to this size.
_RECORD_WIDTH = 400


class FormatError(Exception):
    """Raised when a directory is not a valid xlsx2txt export."""


# ---------------------------------------------------------------------------
# JSON formatting
# ---------------------------------------------------------------------------

def _is_leaf(obj: Any) -> bool:
    values = obj.values() if isinstance(obj, dict) else obj
    return all(not isinstance(v, (dict, list)) for v in values)


def _format(obj: Any, indent: int) -> str:
    compact = json.dumps(obj, ensure_ascii=False)
    if not isinstance(obj, (dict, list)) or not obj:
        return compact
    if indent + len(compact) <= _LINE_WIDTH:
        return compact
    if isinstance(obj, dict) and _is_leaf(obj) and len(compact) <= _RECORD_WIDTH:
        return compact
    pad = " " * (indent + 2)
    end_pad = " " * indent
    if isinstance(obj, dict):
        items = [
            f"{pad}{json.dumps(k, ensure_ascii=False)}: {_format(v, indent + 2)}"
            for k, v in obj.items()
        ]
        return "{\n" + ",\n".join(items) + "\n" + end_pad + "}"
    items = [f"{pad}{_format(v, indent + 2)}" for v in obj]
    return "[\n" + ",\n".join(items) + "\n" + end_pad + "]"


def dumps(obj: Any) -> str:
    """Serialize to JSON with one "record" per line for readable diffs.

    Objects holding only scalars (e.g. a single cell) and containers that
    fit in a line are written inline; larger containers are expanded.
    """
    return _format(obj, 0) + "\n"


# ---------------------------------------------------------------------------
# File names
# ---------------------------------------------------------------------------

def sheet_file_names(names: list[str]) -> list[str]:
    """Build unique, filesystem-safe file names for sheets."""
    used = {"_index"}
    result = []
    for name in names:
        base = safe_file_name(name)
        candidate = base
        counter = 2
        while candidate.lower() in used:
            candidate = f"{base}_{counter}"
            counter += 1
        used.add(candidate.lower())
        result.append(f"{candidate}.json")
    return result


# ---------------------------------------------------------------------------
# Writing
# ---------------------------------------------------------------------------

def _prepare_dir(out_dir: Path, force: bool) -> None:
    if out_dir.exists():
        if not out_dir.is_dir():
            raise FileExistsError(f"Not a directory: {out_dir}")
        is_export = (out_dir / MANIFEST).exists()
        if any(out_dir.iterdir()) and not is_export and not force:
            raise FileExistsError(
                f"Directory is not empty and is not an xlsx2txt export: {out_dir} "
                "(use --force to write anyway)"
            )
        for name in MANAGED:
            target = out_dir / name
            if target.is_dir():
                shutil.rmtree(target)
            elif target.exists():
                target.unlink()
    out_dir.mkdir(parents=True, exist_ok=True)


def write_model(model: dict[str, Any], out_dir: str | Path, force: bool = False) -> list[str]:
    """Write a model to a directory. Returns the list of written files."""
    out_dir = Path(out_dir)
    files: dict[str, bytes] = {}

    def put_json(rel: str, obj: Any) -> None:
        files[rel] = dumps(obj).encode("utf-8")

    put_json(MANIFEST, model["manifest"])
    put_json(WORKBOOK, model["workbook"])

    file_names = sheet_file_names([s["name"] for s in model["sheets"]])
    put_json(SHEETS_INDEX, [
        {"name": sheet["name"], "file": file_name}
        for sheet, file_name in zip(model["sheets"], file_names)
    ])
    for sheet, file_name in zip(model["sheets"], file_names):
        put_json(f"{SHEETS_DIR}/{file_name}", sheet)

    styles = model["styles"]
    for part in STYLE_PARTS + ["cellStyles"]:
        put_json(f"{STYLES_DIR}/{part}.json", styles.get(part, []))

    if model.get("theme"):
        files[THEME] = model["theme"].encode("utf-8")

    if model.get("metadata"):
        files[METADATA] = model["metadata"].encode("utf-8")

    for name, content in sorted((model.get("media") or {}).items()):
        files[f"{MEDIA_DIR}/{name}"] = content

    for name, content in sorted((model.get("printerSettings") or {}).items()):
        files[f"{PRINTER_DIR}/{name}"] = content

    for name, xml in sorted((model.get("pivots") or {}).items()):
        files[f"{PIVOTS_DIR}/{name}"] = xml.encode("utf-8")

    for name, content in sorted((model.get("vba") or {}).items()):
        files[f"{VBA_DIR}/{name}"] = content

    for name, code in sorted((model.get("vbaSources") or {}).items()):
        files[f"{VBA_SOURCES_DIR}/{name}"] = code.encode("utf-8")

    checksums = {rel: hashlib.sha256(content).hexdigest() for rel, content in sorted(files.items())}
    put_json(CHECKSUMS, {"algorithm": "sha256", "files": checksums})

    # Only touch the disk once everything has been serialized successfully.
    _prepare_dir(out_dir, force)
    for rel, content in files.items():
        target = out_dir / rel
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_bytes(content)
    return sorted(files)


# ---------------------------------------------------------------------------
# Reading
# ---------------------------------------------------------------------------

def _load_json(path: Path) -> Any:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except FileNotFoundError:
        raise FormatError(f"Missing file: {path}")
    except json.JSONDecodeError as exc:
        raise FormatError(f"Invalid JSON in {path}: {exc}")


def read_model(in_dir: str | Path) -> dict[str, Any]:
    """Read a directory written by :func:`write_model`."""
    in_dir = Path(in_dir)
    if not (in_dir / MANIFEST).exists():
        raise FormatError(f"Not an xlsx2txt directory (no {MANIFEST}): {in_dir}")

    manifest = _load_json(in_dir / MANIFEST)
    if manifest.get("format") != FORMAT_NAME:
        raise FormatError(f"Unknown format: {manifest.get('format')!r}")
    if manifest.get("formatVersion", 0) > FORMAT_VERSION:
        raise FormatError(
            f"Format version {manifest.get('formatVersion')} is newer than supported ({FORMAT_VERSION})"
        )

    workbook = _load_json(in_dir / WORKBOOK)
    index = _load_json(in_dir / SHEETS_INDEX)
    sheets = []
    for entry in index:
        sheet = _load_json(in_dir / SHEETS_DIR / entry["file"])
        if sheet.get("name") != entry["name"]:
            raise FormatError(
                f"Sheet name mismatch in {entry['file']}: {sheet.get('name')!r} != {entry['name']!r}"
            )
        sheets.append(sheet)

    styles = {}
    for part in STYLE_PARTS + ["cellStyles"]:
        styles[part] = _load_json(in_dir / STYLES_DIR / f"{part}.json")

    theme_path = in_dir / THEME
    # Bytes, not read_text(): keep the original line endings of the theme XML.
    theme = theme_path.read_bytes().decode("utf-8") if theme_path.exists() else None
    metadata_path = in_dir / METADATA
    metadata = metadata_path.read_bytes().decode("utf-8") if metadata_path.exists() else None

    vba = {}
    vba_sources = {}
    vba_dir = in_dir / VBA_DIR
    if vba_dir.is_dir():
        for path in sorted(vba_dir.rglob("*")):
            if not path.is_file():
                continue
            rel = path.relative_to(in_dir).as_posix()
            if rel.startswith(VBA_SOURCES_DIR + "/"):
                vba_sources[rel[len(VBA_SOURCES_DIR) + 1:]] = path.read_bytes().decode("utf-8")
            else:
                vba[path.relative_to(vba_dir).as_posix()] = path.read_bytes()

    printer = {}
    printer_dir = in_dir / PRINTER_DIR
    if printer_dir.is_dir():
        for path in sorted(printer_dir.iterdir()):
            if path.is_file():
                printer[path.name] = path.read_bytes()

    pivots = {}
    pivots_dir = in_dir / PIVOTS_DIR
    if pivots_dir.is_dir():
        for path in sorted(pivots_dir.iterdir()):
            if path.is_file():
                pivots[path.name] = path.read_bytes().decode("utf-8")

    media = {}
    media_dir = in_dir / MEDIA_DIR
    if media_dir.is_dir():
        for path in sorted(media_dir.iterdir()):
            if path.is_file():
                media[path.name] = path.read_bytes()

    return {
        "manifest": manifest,
        "workbook": workbook,
        "styles": styles,
        "sheets": sheets,
        "theme": theme,
        "vba": vba,
        "vbaSources": vba_sources,
        "media": media,
        "pivots": pivots,
        "printerSettings": printer,
        "metadata": metadata,
    }


def check_checksums(in_dir: str | Path) -> dict[str, list[str]]:
    """Compare files against ``_verify/checksums.json``.

    Returns a dict with ``modified``, ``missing`` and ``extra`` file lists.
    """
    in_dir = Path(in_dir)
    data = _load_json(in_dir / CHECKSUMS)
    expected: dict[str, str] = data.get("files", {})
    result: dict[str, list[str]] = {"modified": [], "missing": [], "extra": []}

    for rel, digest in sorted(expected.items()):
        path = in_dir / rel
        if not path.exists():
            result["missing"].append(rel)
        elif hashlib.sha256(path.read_bytes()).hexdigest() != digest:
            result["modified"].append(rel)

    actual = set()
    for top in ("data", VBA_DIR):
        base = in_dir / top
        if base.is_dir():
            actual.update(p.relative_to(in_dir).as_posix() for p in base.rglob("*") if p.is_file())
    if (in_dir / MANIFEST).exists():
        actual.add(MANIFEST)
    result["extra"] = sorted(actual - set(expected))
    return result
