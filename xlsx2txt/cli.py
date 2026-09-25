"""Command-line interface for xlsx2txt."""

import sys
import zipfile
from pathlib import Path

# Allow running as script: python cli.py or .\cli.py
if __name__ == "__main__" and __package__ is None:
    sys.path.insert(0, str(Path(__file__).parent.parent))

import click

from xlsx2txt import __version__
from xlsx2txt.converter import EXCEL_SUFFIXES, export_xlsx, import_dir, load_model, verify_dir
from xlsx2txt.compare import diff_models
from xlsx2txt.storage import FormatError, read_model
from xlsx2txt.text import render_text

# Errors reported without a traceback: broken input files make openpyxl and
# zipfile raise a wide range of exceptions.
INPUT_ERRORS = (FormatError, OSError, ValueError, KeyError, zipfile.BadZipFile)


def _fail(message: str) -> None:
    click.echo(f"Error: {message}", err=True)
    sys.exit(1)


@click.group()
@click.version_option(version=__version__, prog_name="xlsx2txt")
def main():
    """xlsx2txt - Bidirectional Excel converter for Git version control."""
    pass


@main.command()
@click.argument("input_file", type=click.Path(exists=True, dir_okay=False))
@click.argument("output_dir", type=click.Path(file_okay=False), required=False)
@click.option("--mode", type=click.Choice(["normal", "debug"]), default="normal",
              help="debug: print every written file and extra details.")
@click.option("--no-cached-values", is_flag=True,
              help="Do not store last calculated values of formula cells.")
@click.option("--force", is_flag=True, help="Write into a non-empty directory.")
def export(input_file, output_dir, mode, no_cached_values, force):
    """Export XLSX to text format.

    OUTPUT_DIR defaults to the input file name without extension.
    """
    source = Path(input_file)
    if source.suffix.lower() not in EXCEL_SUFFIXES:
        _fail(f"Unsupported file type: {source.suffix or '(none)'}")
    target = Path(output_dir) if output_dir else source.with_suffix("")
    try:
        model = export_xlsx(source, target, force=force, cached_values=not no_cached_values)
    except (FileExistsError, OSError) as exc:
        _fail(str(exc))
    except Exception as exc:  # openpyxl raises various errors for broken files
        _fail(f"Cannot read {source}: {exc}")

    manifest = model["manifest"]
    click.echo(
        f"Exported {source} -> {target} "
        f"({manifest['sheetCount']} sheet(s), {manifest['cellCount']} cell(s))"
    )
    for warning in manifest["warnings"]:
        click.echo(f"Warning: {warning}", err=True)
    if mode == "debug":
        for path in sorted(p for p in target.rglob("*") if p.is_file()):
            click.echo(f"  {path.relative_to(target).as_posix()}")


@main.command("import")
@click.argument("input_dir", type=click.Path(exists=True, file_okay=False))
@click.argument("output_file", type=click.Path(dir_okay=False), required=False)
@click.option("--force", is_flag=True, help="Overwrite OUTPUT_FILE if it exists.")
def import_cmd(input_dir, output_file, force):
    """Import text format to XLSX.

    OUTPUT_FILE defaults to the directory name plus the original extension.
    """
    source = Path(input_dir)
    try:
        if output_file:
            target = Path(output_file)
        else:
            manifest = read_model(source)["manifest"]
            extension = manifest.get("source", {}).get("type", "xlsx")
            target = source.resolve().with_name(f"{source.resolve().name}.{extension}")
        if target.exists() and not force:
            _fail(f"File exists: {target} (use --force to overwrite)")
        import_dir(source, target)
    except (FormatError, ValueError, OSError) as exc:
        _fail(str(exc))
    click.echo(f"Imported {source} -> {target}")


@main.command()
@click.argument("input_dir", type=click.Path(exists=True, file_okay=False))
@click.option("--against", type=click.Path(exists=True, dir_okay=False),
              help="Also compare the data with this Excel file.")
@click.option("--no-roundtrip", is_flag=True, help="Skip the import/export round-trip check.")
def verify(input_dir, against, no_roundtrip):
    """Verify text format integrity."""
    try:
        report = verify_dir(input_dir, against=against, roundtrip=not no_roundtrip)
    except INPUT_ERRORS as exc:
        _fail(str(exc))

    checks = report["checksums"]
    for rel in checks["modified"]:
        click.echo(f"MODIFIED  {rel}")
    for rel in checks["missing"]:
        click.echo(f"MISSING   {rel}")
    for rel in checks["extra"]:
        click.echo(f"EXTRA     {rel} (not referenced by checksums)")
    for error in report["errors"]:
        click.echo(f"ERROR     {error}")
    for line in report["roundtrip"]:
        click.echo(f"ROUNDTRIP {line}")
    for line in report["against"]:
        click.echo(f"MISMATCH  {line}")

    if report["ok"]:
        click.echo("OK")
    else:
        sys.exit(1)


@main.command()
@click.argument("path1", type=click.Path(exists=True))
@click.argument("path2", type=click.Path(exists=True))
@click.option("--ignore-cached", is_flag=True, help="Ignore cached values of formula cells.")
def diff(path1, path2, ignore_cached):
    """Compare two xlsx2txt directories or Excel files.

    Exit code is 0 when identical and 1 when differences were found.
    """
    try:
        changes = diff_models(load_model(path1), load_model(path2), ignore_cached=ignore_cached)
    except INPUT_ERRORS as exc:
        _fail(str(exc))
    for line in changes:
        click.echo(line)
    if changes:
        sys.exit(1)
    click.echo("No differences")


@main.command()
@click.argument("input_path", type=click.Path(exists=True))
@click.option("--no-styles", is_flag=True, help="Do not describe cell formatting.")
@click.option("--no-cached-values", is_flag=True, help="Do not show last calculated formula results.")
def cat(input_path, no_styles, no_cached_values):
    """Print an Excel file or export as plain text, one line per cell.

    Suitable as a Git textconv driver, so that `git diff` shows changes
    inside .xlsx files (see README, "Git integration").
    """
    try:
        model = load_model(input_path, cached_values=not no_cached_values)
    except INPUT_ERRORS as exc:
        _fail(str(exc))
    click.echo(render_text(model, styles=not no_styles), nl=False)


@main.command()
@click.argument("input_path", type=click.Path(exists=True))
def info(input_path):
    """Show information about file or directory."""
    try:
        model = load_model(input_path)
    except INPUT_ERRORS as exc:
        _fail(str(exc))

    manifest = model["manifest"]
    workbook = model["workbook"]
    click.echo(f"Source:   {manifest['source']['name']} ({manifest['source']['type']})")
    click.echo(f"Format:   {manifest['format']} v{manifest['formatVersion']} ({manifest['generator']})")
    props = workbook.get("properties", {})
    labels = {"title": "Title", "creator": "Author", "lastModifiedBy": "Editor",
              "created": "Created", "modified": "Modified"}
    for key, label in labels.items():
        if props.get(key):
            click.echo(f"{label + ':':<9} {props[key]}")
    click.echo(f"Styles:   {len(model['styles'].get('cellStyles', []))}")
    click.echo(f"Names:    {len(workbook.get('definedNames', []))}")
    vba = "no"
    if model.get("vba"):
        modules = len(model.get("vbaSources") or {})
        vba = f"yes ({modules} module(s) in vba/modules)" if modules else "yes"
    click.echo(f"VBA:      {vba}")
    links = workbook.get("externalLinks", [])
    if links:
        click.echo(f"Links:    {', '.join(link['target'] for link in links)}")
    click.echo(f"Sheets:   {len(model['sheets'])}")
    for sheet in model["sheets"]:
        cells = sheet.get("cells", {})
        formulas = sum(1 for c in cells.values() if c.get("t") == "f")
        state = f" [{sheet['state']}]" if sheet.get("state") else ""
        click.echo(
            f"  - {sheet['name']}{state}: range {sheet.get('usedRange')}, "
            f"{len(cells)} cell(s), {formulas} formula(s), "
            f"{len(sheet.get('mergedCells', []))} merged range(s), "
            f"{len(sheet.get('images', []))} image(s), {len(sheet.get('charts', []))} chart(s)"
        )
    for chartsheet in workbook.get("chartsheets", []):
        click.echo(f"  - {chartsheet['name']} (chart sheet): {len(chartsheet.get('charts', []))} chart(s)")
    for warning in manifest.get("warnings", []):
        click.echo(f"Warning: {warning}")


if __name__ == "__main__":
    main()
