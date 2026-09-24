"""High level operations: export, import, verify, diff."""

import tempfile
from pathlib import Path
from typing import Any, Dict, List, Optional, Union

from xlsx2txt.compare import diff_models, validate_model
from xlsx2txt.exporter import export_model
from xlsx2txt.importer import import_model
from xlsx2txt.storage import check_checksums, read_model, write_model

EXCEL_SUFFIXES = {".xlsx", ".xlsm", ".xltx", ".xltm"}

PathLike = Union[str, Path]


def export_xlsx(input_file: PathLike, output_dir: PathLike, force: bool = False,
                cached_values: bool = True) -> Dict[str, Any]:
    """Export an Excel file to an xlsx2txt directory. Returns the model."""
    model = export_model(input_file, cached_values=cached_values)
    write_model(model, output_dir, force=force)
    return model


def import_dir(input_dir: PathLike, output_file: PathLike) -> Path:
    """Build an Excel file from an xlsx2txt directory."""
    model = read_model(input_dir)
    errors = validate_model(model)
    if errors:
        raise ValueError("Invalid xlsx2txt data:\n  " + "\n  ".join(errors))
    return import_model(model, output_file)


def load_model(path: PathLike, cached_values: bool = True) -> Dict[str, Any]:
    """Load a model from an Excel file or an xlsx2txt directory."""
    path = Path(path)
    if path.is_dir():
        return read_model(path)
    return export_model(path, cached_values=cached_values)


def roundtrip_diff(model: Dict[str, Any]) -> List[str]:
    """Import a model into a temporary workbook, re-export it and compare."""
    suffix = ".xlsm" if model.get("vba") else ".xlsx"
    with tempfile.TemporaryDirectory() as tmp:
        target = Path(tmp) / f"roundtrip{suffix}"
        import_model(model, target)
        again = export_model(target, cached_values=False)
    return diff_models(model, again, ignore_cached=True)


def verify_dir(input_dir: PathLike, against: Optional[PathLike] = None,
               roundtrip: bool = True) -> Dict[str, Any]:
    """Verify an xlsx2txt directory.

    Checks file checksums, internal consistency, optionally that the data
    survives an import/export round-trip and that it matches an Excel file.
    """
    input_dir = Path(input_dir)
    model = read_model(input_dir)
    report: Dict[str, Any] = {
        "checksums": check_checksums(input_dir),
        "errors": [],
        "roundtrip": [],
        "against": [],
    }
    report["errors"] = validate_model(model)
    if roundtrip and not report["errors"]:
        report["roundtrip"] = roundtrip_diff(model)
    if against is not None:
        report["against"] = diff_models(model, export_model(against), ignore_cached=True)
    checks = report["checksums"]
    report["ok"] = not (
        checks["modified"] or checks["missing"] or report["errors"]
        or report["roundtrip"] or report["against"]
    )
    return report
