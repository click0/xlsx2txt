# Changelog

All notable changes to xlsx2txt. The section of a version is used as the
text of its GitHub release (see `scripts/release_notes.py`).
Ukrainian version: [CHANGELOG_UK.md](CHANGELOG_UK.md).

## [Unreleased]

### Added
- Shapes: text boxes, rectangles and other preset shapes, connector arrows
  and groups are exported with the sheet (`shapes`: name, text and the
  drawing XML) and restored on import, also on sheets without pictures or
  charts. `cat`, `info` and `diff` show them. Shapes filled with a picture or
  carrying a hyperlink are still only reported in the warnings.
- Example `08-shapes`.

## [0.4.0] - 2026-09-25

**Python 3.11 or newer is now required.** On Python 3.9/3.10 keep using 0.3.0.

### Added
- Pivot tables: the table definition, its cache and cached records are
  exported to `data/pivots/` as indented XML and restored on import; caches
  shared by several pivot tables are stored once. `cat` and `info` list them.
- Example `07-pivot-table`.
- Printer driver settings of sheets (`printerSettings*.bin`) are exported to
  `data/printer/` and restored byte for byte.
- Warnings for everything that is not exported: shapes, SmartArt, threaded
  comments, sparklines, slicers, timelines, form/ActiveX controls, embedded
  objects, data connections, Power Query, the data model, pictures in cells,
  extended conditional formatting/validation and any unknown part of the
  file. openpyxl's own notes are collected into the same list instead of
  being printed to the console.

### Changed
- Python 3.11 or newer is required (was 3.9); tests run on 3.11–3.14.

## [0.3.0] - 2026-09-25

### Added
- `xlsx2txt cat` prints a workbook or an export as plain text, one line per
  cell (values, formulas, formatting, links, comments, VBA code). Used as a
  Git textconv driver, it makes `git diff` show changes inside `.xlsx` files.
- Chart sheets are exported and restored in their original position.
- Links to other workbooks (`[1]Sheet!A1` in formulas) are kept together with
  their cached values.
- Custom document properties (File → Info → Properties → Custom).
- `examples/`: six example workbooks, one per group of features, each with
  its export next to it.

### Fixed
- The workbook theme was reported as changed after reading an export back
  (CRLF line endings were lost).
- Hidden columns with width 0 lost their width.
- Macros are detected by content, so files without an Excel extension
  (for example Git's temporary files) work.

### Changed
- Installation instructions: the package is installed from GitHub releases,
  it is not published on PyPI.
- Excel files in the repository root are ignored by Git.
- Release descriptions are taken from this changelog (English and Ukrainian)
  instead of a generated list of pull requests.

## [0.2.0] - 2026-09-25

### Added
- Images are exported to `data/media/` (one file per distinct image) with
  readable anchors, and restored on import.
- Charts on worksheets (bar, line, pie, …).
- Rich text: cells with several differently formatted runs.
- VBA source code is extracted to `vba/modules/` for review and diffs
  (optional, `pip install xlsx2txt[vba]`); macros are still restored from
  `vbaProject.bin`.
- `info` shows the number of images and charts.

### Changed
- Pillow is now a dependency (needed to write images back).

## [0.1.0] - 2026-09-25

First working version.

### Added
- `export`, `import`, `verify`, `diff` and `info` commands with a full
  xlsx → text → xlsx round-trip.
- Cell values of every type, formulas (including array formulas) and their
  last calculated values, styles, merged cells, row and column sizes,
  frozen panes, hidden sheets, auto filter, print settings, sheet
  protection, data validation, conditional formatting, tables, named ranges,
  comments, hyperlinks, workbook theme and VBA (`vbaProject.bin`).
- One cell per line in the JSON files for readable Git diffs; SHA-256
  checksums in `_verify/`.
- Releases from the GitHub web interface.

[Unreleased]: https://github.com/click0/xlsx2txt/compare/v0.4.0...HEAD
[0.4.0]: https://github.com/click0/xlsx2txt/compare/v0.3.0...v0.4.0
[0.3.0]: https://github.com/click0/xlsx2txt/compare/v0.2.0...v0.3.0
[0.2.0]: https://github.com/click0/xlsx2txt/compare/v0.1.0...v0.2.0
[0.1.0]: https://github.com/click0/xlsx2txt/releases/tag/v0.1.0
