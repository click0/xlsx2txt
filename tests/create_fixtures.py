"""Generate test fixture Excel files."""

from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment


def create_simple_xlsx(path: Path):
    """Create simple.xlsx with plain text in A1."""
    wb = Workbook()
    sheet = wb.active
    sheet["A1"] = "Hello"
    wb.save(path)


def create_styled_xlsx(path: Path):
    """Create styled.xlsx with formatted A1."""
    wb = Workbook()
    sheet = wb.active
    sheet["A1"] = "Styled"

    sheet["A1"].font = Font(
        name="Arial",
        size=14,
        bold=True,
        italic=False,
        color="FF0000",
    )
    sheet["A1"].fill = PatternFill(
        patternType="solid",
        fgColor="FFFF00",
    )
    sheet["A1"].border = Border(
        left=Side(style="thin", color="000000"),
        right=Side(style="thin", color="000000"),
        top=Side(style="thin", color="000000"),
        bottom=Side(style="thin", color="000000"),
    )
    sheet["A1"].alignment = Alignment(
        horizontal="center",
        vertical="center",
    )

    wb.save(path)


def create_empty_xlsx(path: Path):
    """Create empty.xlsx with empty A1."""
    wb = Workbook()
    sheet = wb.active
    # A1 is empty by default
    wb.save(path)


def create_formula_xlsx(path: Path):
    """Create formula.xlsx with formula in A1."""
    wb = Workbook()
    sheet = wb.active
    sheet["A1"] = 10
    sheet["A2"] = 20
    sheet["A3"] = "=A1+A2"
    wb.save(path)


def main():
    fixtures_dir = Path(__file__).parent / "fixtures"
    fixtures_dir.mkdir(exist_ok=True)

    create_simple_xlsx(fixtures_dir / "simple.xlsx")
    create_styled_xlsx(fixtures_dir / "styled.xlsx")
    create_empty_xlsx(fixtures_dir / "empty.xlsx")
    create_formula_xlsx(fixtures_dir / "formula.xlsx")

    print(f"Created test fixtures in {fixtures_dir}")


if __name__ == "__main__":
    main()
