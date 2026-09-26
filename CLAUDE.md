# CLAUDE.md

Guidance for Claude Code working in this repository.

## Project

xlsx2txt — bidirectional converter between Excel (.xlsx/.xlsm) and a
Git-friendly JSON directory layout, with full round-trip.

- `xlsx2txt/exporter.py` — xlsx → model (plain dict)
- `xlsx2txt/importer.py` — model → xlsx
- `xlsx2txt/storage.py` — model ↔ directory (`manifest.json`, `data/`, `vba/`, `_verify/`)
- `xlsx2txt/styles.py`, `drawings.py`, `shapes.py`, `pivots.py`, `vba.py` — styles, images/charts, shapes, pivot tables, VBA sources
- `xlsx2txt/extensions.py`, `threads.py` — worksheet `<extLst>` (sparklines, x14 rules), `<ignoredErrors>`, dynamic array cell metadata; threaded comments
- `xlsx2txt/xmlfrag.py` — XML fragments kept byte for byte (split children, inherited namespaces)
- `xlsx2txt/package.py` — direct zip work: printer settings, shapes, extensions, threaded comments, warnings about parts that are not exported
- `xlsx2txt/compare.py` — semantic diff and validation
- `xlsx2txt/text.py` — plain-text rendering for `cat` / git textconv
- `xlsx2txt/cli.py` — `export`, `import`, `verify`, `diff`, `info`, `cat`

## Checks

```bash
pip install -e ".[dev]"
pytest -q
```

Example workbooks live in `examples/` (built by `examples/generate.py`, with
their exports next to them); test-only data lives in `tests/fixtures/`. After
changing the export format or the examples, run `python examples/generate.py`
and commit the result — `tests/test_examples.py` fails otherwise. Never commit
users' own Excel files.

Python 3.11 or newer is required (`requires-python = ">=3.11"`); CI
(`.github/workflows/tests.yml`) runs the tests on Python 3.11–3.14. Use syntax
available in 3.11 (`list[str]`, `X | None` are fine; no `type X = ...`
statements or other 3.12+ features).

## Workflow

- **Merge your own PRs into `main` automatically** as soon as CI is green on
  the latest commit and there is no merge conflict — do not wait for
  confirmation. If CI fails, fix it first and merge only after it is green.
- Record every user-visible change under `## [Unreleased]` in both
  `CHANGELOG.md` and `CHANGELOG_UK.md`. Before a release, bump `__version__`
  and rename that section to `## [X.Y.Z] - date` in both files: the release
  description is built from it (`scripts/release_notes.py`) and the release
  workflow fails without it.
- The package version lives only in `xlsx2txt/__init__.py` (`__version__`).
- Releases are made from the GitHub web UI (`.github/workflows/release.yml`);
  the release tag must match `__version__`.
- Keep `README.md` and `README_UK.md` in sync when changing documented behavior.
