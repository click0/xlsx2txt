# CLAUDE.md

Guidance for Claude Code working in this repository.

## Project

xlsx2txt — bidirectional converter between Excel (.xlsx/.xlsm) and a
Git-friendly JSON directory layout, with full round-trip.

- `xlsx2txt/exporter.py` — xlsx → model (plain dict)
- `xlsx2txt/importer.py` — model → xlsx
- `xlsx2txt/storage.py` — model ↔ directory (`manifest.json`, `data/`, `vba/`, `_verify/`)
- `xlsx2txt/styles.py`, `drawings.py`, `vba.py` — styles, images/charts, VBA sources
- `xlsx2txt/compare.py` — semantic diff and validation
- `xlsx2txt/text.py` — plain-text rendering for `cat` / git textconv
- `xlsx2txt/cli.py` — `export`, `import`, `verify`, `diff`, `info`, `cat`

## Checks

```bash
pip install -e ".[dev]"
pytest -q
```

CI (`.github/workflows/tests.yml`) runs the tests on Python 3.9–3.13; keep
code compatible with Python 3.9.

## Workflow

- **Merge your own PRs into `main` automatically** as soon as CI is green on
  the latest commit and there is no merge conflict — do not wait for
  confirmation. If CI fails, fix it first and merge only after it is green.
- The package version lives only in `xlsx2txt/__init__.py` (`__version__`).
- Releases are made from the GitHub web UI (`.github/workflows/release.yml`);
  the release tag must match `__version__`.
- Keep `README.md` and `README_UK.md` in sync when changing documented behavior.
