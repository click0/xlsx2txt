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
| Cell values & formulas | ✅ |
| Styles (fonts, colors, borders) | ✅ |
| Merged cells | ✅ |
| Named ranges | ✅ |
| Comments | ✅ |
| Images | ✅ |
| Charts | ✅ |
| VBA macros (.xlsm) | ✅ |
| Conditional formatting | ✅ |
| Data validation | ✅ |

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

# Verify integrity
xlsx2txt verify ./report/
```

## Output Structure

```
report/
├── manifest.json           # Metadata
├── data/
│   ├── workbook.json       # Workbook properties
│   ├── sheets/
│   │   ├── _index.json
│   │   └── Sheet1.json     # Cells, formulas, dimensions
│   └── styles/
│       ├── fonts.json
│       ├── fills.json
│       └── ...
├── vba/                    # VBA code (for .xlsm)
│   └── modules/*.bas
└── _verify/                # Verification data
    └── checksums.json
```

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
