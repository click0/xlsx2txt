# xlsx2txt

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
| VBA macros (.xlsm) | ✅ (binary `vbaProject.bin` preserved) |
| Images | 🚧 (reported as a warning) |
| Charts | 🚧 (reported as a warning) |

## Installation

```bash
pip install xlsx2txt
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

## Commands

| Command | Description |
|---------|-------------|
| `export FILE [DIR]` | Export `.xlsx`/`.xlsm` into a directory (default: file name without extension). `--no-cached-values` skips last calculated formula results, `--force` writes into a non-empty foreign directory, `--mode debug` lists written files. |
| `import DIR [FILE]` | Build an Excel file (default: directory name + original extension). Refuses to overwrite an existing file without `--force`. |
| `verify DIR` | Checks checksums, internal consistency and that the data survives an import → export round-trip (`--no-roundtrip` to skip). `--against FILE` also compares with an Excel file. Exit code 1 on problems. |
| `diff A B` | Semantic diff of two exports or Excel files (styles are compared by content, not by index). `--ignore-cached` ignores cached formula results. Exit code 1 when different. |
| `info PATH` | Sheets, cell/formula counts, VBA presence and unsupported features. |

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
│   ├── workbook.json       # Properties, sheet order, defined names, epoch
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
│   └── theme/theme1.xml    # Workbook theme (theme colors)
├── vba/                    # VBA project (for .xlsm)
│   └── xl/vbaProject.bin
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
`d` date/time in ISO 8601, `td` duration in seconds, `e` error, `f` formula),
`f` formula, `s` style index in `cellStyles.json` (omitted for the default
style), `link` hyperlink, `comment` comment. For formula cells `v` is the last
value calculated by Excel; it is informational and ignored on import.

## Limitations

- Images, charts, pivot tables, chartsheets and external links are not exported
  yet; `export` and `info` print a warning when a file contains them.
- VBA code is kept as the binary `vbaProject.bin`, not as `.bas` sources.
- Rich text inside a cell is stored as plain text.
- Formula results are not recalculated; Excel recalculates them when the
  restored file is opened.

## Comparison

| Tool | Direction | Formulas | Styles | Round-trip |
|------|-----------|----------|--------|------------|
| **xlsx2txt** | ↔️ bidirectional | ✅ | ✅ | ✅ |
| xlsx2csv | → one-way | ❌ | ❌ | ❌ |
| xlsxgrep | → one-way | ❌ | ❌ | ❌ |

## License

BSD 3-Clause License

## Copyright

Copyright (c) 2025, Vladyslav V. Prodan

## Acknowledgments

Inspired by the need for proper Excel version control.
Name used with permission from [@cloudwu](https://github.com/cloudwu/xlsx2txt).
