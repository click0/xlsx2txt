"""Command-line interface for xlsx2txt."""

import sys
from pathlib import Path

# Allow running as script: python cli.py or .\cli.py
if __name__ == "__main__" and __package__ is None:
    sys.path.insert(0, str(Path(__file__).parent.parent))

import click

from xlsx2txt import __version__


@click.group()
@click.version_option(version=__version__, prog_name="xlsx2txt")
def main():
    """xlsx2txt - Bidirectional Excel converter for Git version control."""
    pass


@main.command()
@click.argument("input_file", type=click.Path(exists=True))
@click.argument("output_dir", type=click.Path(), required=False)
@click.option("--mode", type=click.Choice(["normal", "debug"]), default="normal")
def export(input_file, output_dir, mode):
    """Export XLSX to text format."""
    click.echo(f"Export: {input_file} -> {output_dir or 'auto'} (mode={mode})")
    click.echo("Not implemented yet.")


@main.command("import")
@click.argument("input_dir", type=click.Path(exists=True))
@click.argument("output_file", type=click.Path(), required=False)
def import_cmd(input_dir, output_file):
    """Import text format to XLSX."""
    click.echo(f"Import: {input_dir} -> {output_file or 'auto'}")
    click.echo("Not implemented yet.")


@main.command()
@click.argument("input_dir", type=click.Path(exists=True))
def verify(input_dir):
    """Verify text format integrity."""
    click.echo(f"Verify: {input_dir}")
    click.echo("Not implemented yet.")


@main.command()
@click.argument("path1", type=click.Path(exists=True))
@click.argument("path2", type=click.Path(exists=True))
def diff(path1, path2):
    """Compare two xlsx2txt directories."""
    click.echo(f"Diff: {path1} vs {path2}")
    click.echo("Not implemented yet.")


@main.command()
@click.argument("input_path", type=click.Path(exists=True))
def info(input_path):
    """Show information about file or directory."""
    click.echo(f"Info: {input_path}")
    click.echo("Not implemented yet.")


if __name__ == "__main__":
    main()
