# xlsx2txt

[Українською](https://github.com/click0/xlsx2txt/blob/main/README_UK.md)

**Bidirectional converter between Excel (.xlsx/.xlsm) and text-based JSON format with full round-trip support.**

Unlike simple text extractors, xlsx2txt preserves everything — formulas, styles, merged cells, VBA macros — and can fully restore the original Excel file.

## Why?

- 📊 **Git-friendly** — meaningful diffs for Excel files
- 🔄 **Round-trip** — xlsx → text → xlsx without data loss
- 👁️ **Human-readable** — JSON format for easy review
- ✅ **Verifiable** — built-in integrity checks

## Features

| Feature | Status |
|---------|--------|
| Cell values & formulas (incl. array formulas) | ✅ |
| Styles (fonts, fills, borders, alignment, number formats, protection) | ✅ |
| Merged cells | ✅ |
| Named ranges (workbook and sheet scope) | ✅ |
| Comments | ✅ |
| Hyperlinks | ✅ |
| Column widths / row heights, hidden & outline levels | ✅ |
| Freeze panes, tab colors, hidden sheets | ✅ |
| Auto filter, print settings, sheet protection | ✅ |
| Tables (ListObjects) | ✅ |
| Conditional formatting | ✅ |
| Data validation | ✅ |
| VBA macros (.xlsm) | ✅ (binary `vbaProject.bin` preserved, source code in `vba/modules/` for review) |
| Rich text (formatted runs inside a cell) | ✅ |
| Images | ✅ (stored in `data/media/`) |
| Charts | ✅ (chart XML + anchor) |
| Chart sheets | ✅ |
| Shapes (text boxes, arrows, connectors, groups) | ✅ (drawing XML in the sheet file) |
| Sparklines, extended conditional formatting and data validation (Excel 2010+) | ✅ (`extensions` in the sheet file) |
| Threaded comments (Excel 365 conversations) | ✅ (`threadedComments` in the sheet file, authors in `workbook.json`) |
| Pivot tables | ✅ (definition and cache as indented XML in `data/pivots/`) |
| Printer driver settings | ✅ (`data/printer/`, restored byte for byte) |
| External links to other workbooks | ✅ |

## Installation

Requires Python 3.11 or newer. The package is not published on PyPI yet.
Install it from GitHub:

```bash
# latest release (replace the version with the one you need)
pip install https://github.com/click0/xlsx2txt/releases/download/v0.5.0/xlsx2txt-0.5.0-py3-none-any.whl

# or the current main branch
pip install git+https://github.com/click0/xlsx2txt.git

# with VBA source extraction
pip install "xlsx2txt[vba] @ git+https://github.com/click0/xlsx2txt.git"
```

## Quick Start

```bash
# Export Excel to text format
xlsx2txt export report.xlsx ./report/

# Edit JSON files or commit to Git
git add report/
git commit -m "Q1 2025 updates"

# Restore Excel file
xlsx2txt import ./report/ report_new.xlsx

# Verify integrity (checksums + import/export round-trip)
xlsx2txt verify ./report/
xlsx2txt verify ./report/ --against report.xlsx

# Compare two exports and/or Excel files
xlsx2txt diff report.xlsx report_new.xlsx

# Summary of a file or directory
xlsx2txt info report.xlsx
```

## Examples

The [`examples/`](https://github.com/click0/xlsx2txt/tree/main/examples/) directory contains small workbooks for every group of
features together with their exports, e.g. [`examples/02-styles/`](https://github.com/click0/xlsx2txt/tree/main/examples/02-styles/).

## Commands

| Command | Description |
|---------|-------------|
| `export FILE [DIR]` | Export `.xlsx`/`.xlsm` into a directory (default: file name without extension). `--no-cached-values` skips last calculated formula results, `--force` writes into a non-empty foreign directory, `--mode debug` lists written files. |
| `import DIR [FILE]` | Build an Excel file (default: directory name + original extension). Refuses to overwrite an existing file without `--force`. |
| `verify DIR` | Checks checksums, internal consistency and that the data survives an import → export round-trip (`--no-roundtrip` to skip). `--against FILE` also compares with an Excel file. Exit code 1 on problems. |
| `diff A B` | Semantic diff of two exports or Excel files (styles are compared by content, not by index). `--ignore-cached` ignores cached formula results. Exit code 1 when different. |
| `info PATH` | Sheets, cell/formula counts, VBA presence and unsupported features. |
| `cat PATH` | Print an Excel file or export as plain text, one line per cell (`--no-styles`, `--no-cached-values`). Used for `git diff`, see below. |

The same operations are available from Python:

```python
from xlsx2txt import export_xlsx, import_dir, verify_dir, load_model, diff_models

export_xlsx("report.xlsx", "report/")
import_dir("report/", "report_new.xlsx")
print(diff_models(load_model("report.xlsx"), load_model("report_new.xlsx")))
```

## Output Structure

```
report/
├── manifest.json           # Format version, source file, warnings
├── data/
│   ├── workbook.json       # Properties, sheet order, defined names, epoch, comment authors
│   ├── sheets/
│   │   ├── _index.json     # Sheet name -> file name
│   │   └── Sheet1.json     # Cells, formulas, dimensions, merges, rules
│   ├── styles/
│   │   ├── fonts.json
│   │   ├── fills.json
│   │   ├── borders.json
│   │   ├── alignments.json
│   │   ├── protections.json
│   │   └── cellStyles.json # Combinations referenced by cells ("s")
│   ├── theme/theme1.xml    # Workbook theme (theme colors)
│   ├── media/              # Images, named by content hash
│   ├── pivots/             # Pivot tables and their caches (XML)
│   └── printer/            # Printer driver settings of sheets (binary)
├── vba/                    # VBA project (for .xlsm)
│   ├── xl/vbaProject.bin
│   └── modules/*.bas       # VBA source code (read-only, needs oletools)
└── _verify/                # Verification data
    └── checksums.json
```

Every cell is written on its own line, so Git diffs stay readable:

```json
"cells": {
  "A1": {"v": "Sales Report Q1 2025", "t": "s", "s": 1},
  "B3": {"v": 100, "t": "n"},
  "D3": {"t": "f", "f": "=B3*C3", "v": 999},
  "E3": {"v": "2025-03-14T00:00:00", "t": "d", "s": 3}
}
```

Cell keys: `v` value, `t` type (`s` string, `n` number, `b` boolean,
`rs` rich text, `d` date/time in ISO 8601, `td` duration in seconds, `e` error, `f` formula),
`f` formula, `s` style index in `cellStyles.json` (omitted for the default
style), `link` hyperlink, `comment` comment. For formula cells `v` is the last
value calculated by Excel; it is informational and ignored on import.

Shapes are listed in the sheet file under `shapes`: `name` and `text` are a
preview for diffs, `xml` is the shape as Excel stores it and is what import
uses — edit the text inside `xml`. `rels` holds what the shape links to (a
hyperlink target, or a picture fill stored in `data/media/`), `layer` the
number of pictures and charts drawn below it (absent when it is on top).

`extensions` keeps the sheet's Excel 2010+ extensions (sparklines, extended
conditional formatting such as negative data bars, drop-down lists from
other sheets) as XML; a conditional formatting rule linked to one has
`extId`. `threadedComments` has one record per comment or reply: `ref` cell,
`personId` (see `persons` in `workbook.json`), `dT` date, `text`, `parentId`
for replies, `done` for resolved threads; `xml` keeps mentions.

## Git integration

Let `git diff` show what changed inside `.xlsx` files, without exporting them:

```bash
echo '*.xlsx diff=xlsx' >> .gitattributes
echo '*.xlsm diff=xlsx' >> .gitattributes
git config diff.xlsx.textconv "xlsx2txt cat"
```

```diff
-B3: 100
+B3: 150
-A4: 'Gadget'
+A4: 'Gadget'  [font bold color FFFF0000]
+A7: 'New row'
```

## Limitations

- Not exported: SmartArt, slicers and timelines, form and ActiveX controls, embedded OLE
  objects, data connections and Power Query, the data model, pictures in
  cells. `export` and `info` warn about each of them (and about any other
  part of the file they do not know), so nothing is lost silently.
- VBA is restored from the binary `vbaProject.bin`. The sources in
  `vba/modules/` are extracted for review and diffs only (install the `vba`
  extra: `pip install xlsx2txt[vba]`); editing them does not change the macros.
- The calculation chain is not kept; Excel recreates it when the file is
  opened. A hidden
  column with width 0 comes back with the default width (it stays hidden).
- Charts are stored as chart XML as understood by openpyxl; exotic chart
  features that openpyxl does not support may be lost.
- The drawing order of shapes relative to pictures and charts is kept; pictures
  and charts among themselves are restored charts first.
- Formula results are not recalculated; Excel recalculates them when the
  restored file is opened.

## Releasing

Releases are made from the GitHub web UI (workflow `.github/workflows/release.yml`):

1. Bump `__version__` in `xlsx2txt/__init__.py` (the package version is read from it), describe the
   version in [CHANGELOG.md](https://github.com/click0/xlsx2txt/blob/main/CHANGELOG.md) and [CHANGELOG_UK.md](https://github.com/click0/xlsx2txt/blob/main/CHANGELOG_UK.md) (section
   `## [X.Y.Z] - date`) and merge to `main`. That section becomes the release description.
2. Either
   - **Actions → release → Run workflow**, enter the version (e.g. `0.1.0`); the workflow runs the tests,
     builds the wheel and sdist, creates the tag `v0.1.0` and a GitHub Release with the files; or
   - **Releases → Draft a new release**, create tag `v0.1.0`, press **Publish**; the workflow runs the tests
     and attaches the built files to that release. Leave the description empty (or use
     "Generate release notes"): it is replaced by the changelog text; a description typed by hand is kept.
3. Optional PyPI publishing (no tokens needed, [trusted publishing](https://docs.pypi.org/trusted-publishers/)):
   - on pypi.org: **Your projects → Publishing → Add a new pending publisher → GitHub**, with
     PyPI project name `xlsx2txt`, owner `click0`, repository `xlsx2txt`, workflow `release.yml`,
     environment `pypi`;
   - then either tick **pypi** when running the workflow, or set the repository variable
     `PUBLISH_TO_PYPI=true` (**Settings → Secrets and variables → Actions → Variables**) so that
     releases published from the Releases page go to PyPI too. The first upload creates the project.

## Comparison

| Tool | Direction | Formulas | Styles | Round-trip |
|------|-----------|----------|--------|------------|
| **xlsx2txt** | ↔️ bidirectional | ✅ | ✅ | ✅ |
| xlsx2csv | → one-way | ❌ | ❌ | ❌ |
| xlsxgrep | → one-way | ❌ | ❌ | ❌ |

## License

BSD 3-Clause License

## Copyright

Copyright (c) 2025-2026, Vladyslav V. Prodan

## Acknowledgments

Inspired by the need for proper Excel version control.
Name used with permission from [@cloudwu](https://github.com/cloudwu/xlsx2txt).
