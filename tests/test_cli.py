"""Tests for the command-line interface."""

from click.testing import CliRunner

from xlsx2txt.cli import main

from tests.conftest import FIXTURES_DIR


def run(*args):
    return CliRunner().invoke(main, [str(a) for a in args])


def test_version():
    result = run("--version")
    assert result.exit_code == 0 and "xlsx2txt" in result.output


def test_export_import_verify(tmp_path):
    out_dir = tmp_path / "complex"
    result = run("export", FIXTURES_DIR / "complex.xlsx", out_dir)
    assert result.exit_code == 0, result.output
    assert "1 sheet(s)" in result.output

    result = run("verify", out_dir, "--against", FIXTURES_DIR / "complex.xlsx")
    assert result.exit_code == 0, result.output
    assert result.output.strip().endswith("OK")

    result = run("import", out_dir)
    assert result.exit_code == 0, result.output
    assert (tmp_path / "complex.xlsx").exists()

    result = run("import", out_dir)
    assert result.exit_code == 1 and "File exists" in result.output

    result = run("diff", FIXTURES_DIR / "complex.xlsx", tmp_path / "complex.xlsx", "--ignore-cached")
    assert result.exit_code == 0, result.output
    assert "No differences" in result.output


def test_export_default_dir(tmp_path):
    source = tmp_path / "simple.xlsx"
    source.write_bytes((FIXTURES_DIR / "simple.xlsx").read_bytes())
    result = run("export", source, "--mode", "debug")
    assert result.exit_code == 0, result.output
    assert (tmp_path / "simple" / "manifest.json").exists()
    assert "data/sheets/Sheet.json" in result.output


def test_diff_exit_code(tmp_path):
    result = run("diff", FIXTURES_DIR / "simple.xlsx", FIXTURES_DIR / "styled.xlsx")
    assert result.exit_code == 1
    assert "A1.v: 'Hello' -> 'Styled'" in result.output


def test_verify_reports_modified(tmp_path):
    out_dir = tmp_path / "simple"
    run("export", FIXTURES_DIR / "simple.xlsx", out_dir)
    sheet = out_dir / "data" / "sheets" / "Sheet.json"
    sheet.write_text(sheet.read_text().replace("Hello", "Bye"))
    result = run("verify", out_dir)
    assert result.exit_code == 1
    assert "MODIFIED  data/sheets/Sheet.json" in result.output


def test_info(tmp_path):
    result = run("info", FIXTURES_DIR / "complex.xlsx")
    assert result.exit_code == 0, result.output
    assert "Data: range A1:D6, 19 cell(s), 4 formula(s), 1 merged range(s)" in result.output


def test_bad_input(tmp_path):
    bogus = tmp_path / "bogus.xlsx"
    bogus.write_text("not a zip")
    result = run("info", bogus)
    assert result.exit_code == 1 and "Error" in result.output

    result = run("export", tmp_path / "bogus.xlsx", tmp_path / "out")
    assert result.exit_code == 1 and "Error" in result.output

    result = run("verify", tmp_path)
    assert result.exit_code == 1 and "Not an xlsx2txt directory" in result.output
