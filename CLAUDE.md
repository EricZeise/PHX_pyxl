# CLAUDE.md — PHPP Data Tool (openpyxl backend)

## What this project is

A Python CLI that **reads designer-entered input data from a filled PHPP workbook** (Passive House Planning Package), **stores it as a portable JSON record**, and **writes that record back into a blank PHPP workbook**.

This is the **openpyxl-only** variant — it does **not require Excel to be installed** for reading or writing. Reading uses openpyxl's dual-load approach: `data_only=True` for cached values and label searching, `data_only=False` for formula detection. Writing resolves addresses via openpyxl (read-only) but persists cell values via a surgical ZIP/XML patch (`surgical_writer.py`, using `lxml`) rather than openpyxl's save — this preserves `<extLst>` extensions (Data Validation, etc.) and `<headerFooter>` content that an openpyxl save would otherwise drop.

### MVP pipeline

```
Filled PHPP (.xlsx)  →  read  →  JSON record  →  write  →  Blank PHPP (.xlsx)
```

### Relationship to PHX_xlwg

PHX_xlwg uses xlwings (requires Excel) for live formula recalculation and cell addressing. PHX_pyxl uses the same field map, models, and map parser but replaces xlwings with openpyxl's dual-load approach. The trade-off: no Excel dependency, but formula results are cached (not recalculated). Both projects persist writes via the same surgical ZIP/XML patch, so file integrity on write is now equivalent between them.

---

## Architecture

```
phpp-field-mapping/EN_10_6_IP.md   (locator dictionary — where each field lives, versioned by PHPP type)
        ↓
   map_parser.py        (parse markdown → structured dict; requires explicit type tags)
        ↓
   locators.py          (6 addressing strategies, openpyxl dual-load)
        ↓
 ┌──────┴──────┐
 reader.py    writer.py  (openpyxl resolution + surgical XML persistence)
 └──────┬──────┘
     models.py           (pydantic validation)
        ↓
     cli.py              (Click CLI, --phpp-version selects the field map)
```

### Versioned field maps

`phpp-field-mapping/` holds one field map per PHPP workbook type — `EN_10_6_IP.md` (IP-shell workbook, which also carries `<Name> SI`-suffixed mirror tabs) and `EN_10_6_SI.md` (genuinely SI-native, single-shell workbook). Selected via `--phpp-version` (default `EN_10_6_IP`). There is no runtime SI/IP sheet-name guessing — each version declares its own correct `sheet_name` directly. This matters because in an IP-shell workbook, the `<Name> SI` tabs are formula mirrors of the base tab's real input cells, not independent data — reading them under `skip_formulas` used to silently lose designer input.

### Dual-load approach

Every read operation loads the workbook twice:
- `data_only=True` — cached values for reading and label searching
- `data_only=False` — formula strings for input/formula classification

Locator functions accept a `WsPair = tuple[Worksheet, Worksheet]` (values sheet, formulas sheet).

---

## Repository structure

```
PHX_pyxl/
├── CLAUDE.md                    ← this file
├── pyproject.toml               ← deps: openpyxl, lxml, click, pydantic (NO xlwings); Ruff / isort / ty / pytest settings
├── uv.lock                      ← locked development environment (`uv sync`)
├── phpp-field-mapping/           ← versioned locator dictionaries (31 worksheets each)
│   ├── EN_10_6_IP.md             ← IP-shell workbook (base tabs + SI mirror tabs)
│   └── EN_10_6_SI.md             ← genuinely SI-native, single-shell workbook
├── src/
│   ├── phpp_tool/
│   │   ├── __init__.py
│   │   ├── cli.py               ← Click CLI: read / write / inspect-map, --phpp-version
│   │   ├── map_parser.py        ← Parse a field map file → structured dict; requires type tags
│   │   ├── locators.py          ← 6 addressing strategies (openpyxl dual-load)
│   │   ├── reader.py            ← openpyxl-based reader
│   │   ├── writer.py            ← Resolves addresses via openpyxl, collects writes
│   │   ├── atomic.py            ← Atomic file output (temp file beside the destination, then replace)
│   │   ├── surgical_writer.py   ← Persists writes via ZIP/XML patch (lxml), preserving extLst/headerFooter
│   │   └── models.py            ← Pydantic models for building record JSON
│   └── compare_json/            ← Standalone JSON diff tool
├── scripts/
│   ├── roundtrip.py             ← Two-part roundtrip test (Parts 1 & 2), accepts --phpp-version
│   └── verify_excel.py          ← Post-Excel full-fidelity comparison
├── tests/
└── records/                     ← Output directory for JSON building records
```

---

## Commands

```bash
# Set up the development environment (uv: pyproject.toml + uv.lock; installs the package editable plus
# pytest, Ruff, isort, ty, xlwings and lxml-stubs). `pip install -e ".[dev]"` still works without uv.
uv sync

# Run tests (no Excel needed — 114 tests, <1 second)
uv run pytest tests/ -v

# The review gate (PH-Demonstrator's CodeReview.md, Step 1) -- all must pass before a commit
uv run ruff check . && uv run ruff format --check . && uv run isort --check-only . \
  && uv run ty check --error-on-warning && uv lock --check && uv run pytest -q

# Step 2 (detailed review): report-only counts may only go down; coverage may not drop
uv run python scripts/check_ratchet.py && uv run python scripts/run_with_coverage.py

# Read a PHPP into JSON (--phpp-version defaults to EN_10_6_IP)
phpp-tool read Data/Example_IP.xlsx -o records/my_building.json --phpp-version EN_10_6_IP

# Read a genuinely SI-native PHPP
phpp-tool read Data/Example_SI.xlsx -o records/my_building.json --phpp-version EN_10_6_SI

# Write a record into a blank PHPP -- must match the --phpp-version used for read
phpp-tool write records/my_building.json Data/Empty_IP.xlsx -o output.xlsx --phpp-version EN_10_6_IP

# Roundtrip test (Part 1: openpyxl only; Part 2: xlwings+Excel if available)
python scripts/roundtrip.py Data/Example_IP.xlsx Data/Empty_IP.xlsx
python scripts/roundtrip.py --phpp-version EN_10_6_SI Data/Example_SI.xlsx Data/Empty_SI.xlsx

# Post-Excel verification (after manually opening both files in Excel and saving)
python scripts/verify_excel.py Data/Example_IP.xlsx records/.../Example_IP_written.xlsx
```

---

## Constraints

- **Never read or embed PHPP formulas** — only designer-entered input values.
- **The field map is the single source of truth** for cell locations.
- **No Excel required** for reading, writing, or Part 1 verification.
- **Writes preserve file integrity** — persistence goes through `surgical_writer.py`'s ZIP/XML patch, not openpyxl's save, so `<extLst>` extensions (Data Validation, etc.) and `<headerFooter>` content survive the round trip. Verified byte-for-byte across all 83 sheets of a full roundtrip write.
- **Cached formula values** — without Excel, formula results reflect the last Excel save, not a live recalculation. Part 2 of the roundtrip test uses xlwings+Excel (optional) to verify recalculation integrity.
- **Read and write must use the same `--phpp-version`** for a given workbook — `read` stamps the output JSON with `_phpp_version`; `write` warns (doesn't hard-block) if it doesn't match the `--phpp-version` passed to `write`.
- **Every config/items field-map entry requires an explicit type tag** (`(literal)`/`(address)`/`(named_range)`) — `map_parser.py` raises `FieldMapError` at parse time if one is missing or invalid. There's no shape-based type inference anymore.

---

## Review gate (2026-11-09)

PH-Demonstrator's Step 1 code-review gate is applied here: `pyproject.toml` carries the tool settings (line length 88, target Python 3.10 = the minimum supported; Ruff rules E, W, F, B, UP, C4, PIE, PLW, PLE; isort profile black; ty at Python 3.10 with a `tests/**` override that relaxes `unresolved-attribute`, `invalid-argument-type` and `not-subscriptable` because the tests read Optional sections of the pydantic record directly) and `uv.lock` the development environment (`.python-version` 3.13). `lxml-stubs` is a dev dependency because `lxml.etree` is compiled and the checker cannot read it. The environment fix that came with it: the old `.venv` pointed its editable install at a path from before the repo moved under `PHX/`, so `phpp_tool` was not importable and pytest could not collect -- `uv sync` rebuilt it. Fixes made: unused imports and f-strings without placeholders removed, four loop variables that were overwritten renamed (`locators.resolve_named_range`, `map_parser._parse_options`, `writer` named-range resolution), 17 files reformatted. Behaviour is unchanged (114 tests). **Atomic output (same day):** every file the package produces -- the written PHPP workbook (`surgical_writer`), the JSON building record (`phpp-tool read -o`) and the `compare_json` reports -- now goes through `phpp_tool/atomic.py` (`atomic_write_bytes` / `atomic_write_text` / `atomic_copy`): a temporary file beside the destination, `fsync`, then `os.replace`, so a crash, a full disk or an interruption leaves the previous output exactly as it was and no half-written file; the temporary file is removed on any failure, an existing file keeps its permissions, and writing over the template itself still works. The template and input workbooks were never modified. `tests/test_atomic.py` covers the failure paths (114 tests).

## Step 2 (2026-11-09, later)

The detailed review of PH-Demonstrator's `CodeReview.md` was run: Ruff also enforces SIM, PERF, FURB, RUF, ARG, S, T20, N and BLE (ignored on purpose: `S101` -- the tests' asserts; `cli.py` and the scripts may print; scripts / tests may run subprocesses). Findings fixed: `if` / `elif` chains and `with` blocks combined (`SIM102`, `SIM114`, `SIM117`), `x if x else None` -> `x or None`, three `append` loops -> `extend`, `[*a, b]` instead of `a + [b]`, two unused private parameters removed (`reader._read_section(ws_config)`, `writer._write_items(ws_labels)`), `airtightness_Vn50` kept (it is the PHPP field's own name, part of the record schema; `noqa: N815` with that reason) and one `noqa: BLE001` for the Excel-launch probe in `scripts/roundtrip.py`. `scripts/check_ratchet.py` holds the report-only `C90` / `PLR` counts (12 complex functions, 75 magic-value comparisons, ...) and the `noqa` / `ty: ignore` counts against `scripts/ratchet_baseline.json`; `scripts/run_with_coverage.py` runs pytest under coverage (7.17.2, Python 3.11+ only: the library itself supports 3.10) against `scripts/coverage_baseline.json` (62.14 % statements + branches -- the reader and writer need real PHPP workbooks, which are not in the repository, so their coverage is low). `uv lock --upgrade --dry-run` found nothing newer and `pip-audit` of the locked dependencies found no known vulnerabilities. The ty `tests/**` override still relaxes three rules (Optional sections of the pydantic record). Not done: the complexity refactors, tests that run against a (synthetic) PHPP workbook, CI, Python 3.10 verification (the tools check against 3.10, but no 3.10 interpreter was run).

