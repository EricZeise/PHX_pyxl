# PHX_pyxl: Strategy and Routine Reference

## Part 1 — Strategy Summary

### Goal

Read designer-entered input data from a filled PHPP workbook (Passive House Planning Package, ~29 MB, 83 sheets), store it as portable JSON, and write that JSON back into a blank PHPP template — preserving every input value while leaving all formulas untouched. Do this entirely with openpyxl, requiring no Excel installation.

### Why openpyxl (and why not xlwings)

The sibling project PHX_xlwg uses xlwings to drive a live Excel instance for reading and writing. This works well on machines where Excel is installed, but introduces hard dependencies: Excel must be present, AppleScript (macOS) or COM (Windows) must be functional, and the process is bound to a GUI application's lifecycle. macOS 26 introduced an AppleScript bug where Excel save operations hang indefinitely on large workbooks, forcing PHX_xlwg into a hybrid architecture.

PHX_pyxl eliminates these dependencies entirely. openpyxl is a pure-Python library that reads and writes .xlsx files as XML archives — no Excel required. The trade-off is that formula results are cached (not recalculated), but for a tool whose purpose is transferring *input* values between workbooks, cached formula results are irrelevant: the tool only reads and writes input cells.

### The dual-load architecture

openpyxl's `data_only` parameter controls what it returns for formula cells:

- `data_only=True` — returns the cached result (the last value Excel computed before saving). Labels and text are readable. Formula cells return their displayed value.
- `data_only=False` — returns the formula string (e.g. `=SUM(A1:A10)`). Labels that are formulas return `=` strings instead of display text.

Neither mode alone suffices. The reader needs cached label text (to search for locator strings) *and* formula strings (to distinguish inputs from formulas). The solution is to load the workbook twice:

| Load | `data_only` | Purpose |
|------|-------------|---------|
| `wb_vals` | `True` | Read cached cell values and search for label text |
| `wb_fmls` | `False` | Detect formula cells (value starts with `=`) |

These two loads are paired throughout the codebase as a `WsPair = tuple[Worksheet, Worksheet]` — the values worksheet and the formulas worksheet for the same sheet.

### Formula-aware filtering

Roughly 35% of mapped cells in a typical PHPP workbook are formula results. Writing a formula's cached result back into a template would overwrite the formula itself, breaking the spreadsheet's calculation chain.

The reader's `skip_formulas` flag (on by default) checks each cell against the `data_only=False` worksheet before reading. If the formula string starts with `=`, the cell is skipped and recorded as `None` in the JSON. This ensures the JSON contains only values that can be meaningfully written back.

### The writer's resolve-then-patch approach

The writer loads the template once, read-only, then persists writes via a separate surgical patch step:

| Step | Mechanism | Purpose |
|------|-----------|---------|
| Resolve | `wb_labels = load_workbook(..., data_only=True)` | Resolve locator lookups — label searches, named ranges, entry-row detection — using cached text values. Never saved; discarded after collecting writes. |
| Persist | `surgical_writer.apply_surgical_writes()` | Edits the `.xlsx` as a ZIP archive, patching only the `<sheetData>` region of sheets with writes via `lxml`. Never invokes openpyxl's serializer, so `<extLst>` extensions and `<headerFooter>` content pass through as untouched original bytes. |

This replaced an earlier design that used a second `data_only=False` openpyxl load to apply writes and `save()` the workbook — that approach preserved formulas in untouched cells correctly, but openpyxl's save cycle also dropped `<extLst>` extensions (Data Validation, etc.) and mangled `<headerFooter>` content it couldn't parse. See Part 5 (Concerns, Features, and Limitations) for the fix's verification.

### The field map

A markdown file per PHPP workbook type — `phpp-field-mapping/EN_10_6_IP.md` and `EN_10_6_SI.md`, selected via `--phpp-version` — serves as the dictionary of every mapped cell across 31 worksheets. It defines six addressing strategies for locating cells, since PHPP's layout varies from sheet to sheet:

1. **Label-anchored** — find a text label in one column, read/write the value in a paired column at an optional row offset.
2. **Header + entry block** — find a header row and entry row, then iterate repeating data rows reading each column field.
3. **Named ranges** — resolve an Excel defined name (often German, e.g. `Werte_Klima_Region`) to its cell.
4. **Absolute address** — a fixed cell reference like `C11`.
5. **Column + row-offset** — locate an anchor row, then read at a column with a fixed row offset.
6. **Fixed result rows/cols** — a fixed row and column, typically for result cells.

### Resolved limitation: openpyxl save degradation (fixed 2026-07-01)

openpyxl's load→save cycle strips some Excel extensions: data validation rules, custom headers/footers, and certain XML namespace extensions that PHPP relies on. This no longer affects written output, because the writer never calls openpyxl's `save()` — `surgical_writer.py` persists writes via a ZIP/XML patch instead (see above). Verified byte-for-byte: `<extLst>` and `<headerFooter>` regions are identical between template and written output across all 83 sheets of a full 13,102-cell roundtrip write.

The remaining caveat is unrelated to this fix: written files still cannot be reopened by Excel via AppleScript automation on macOS 26 (data validation errors cause a hang) — that's a macOS/Excel-version compatibility issue, not something either openpyxl or the surgical patch controls.

### Verification

The roundtrip test confirms data fidelity end to end. The writer returns its full list of `(sheet_name, col, row, value)` writes. The test script opens the written file with openpyxl and verifies every cell matches — 13,102 cells, zero mismatches. An optional second part uses xlwings+Excel to validate that openpyxl's cached formula values match Excel's live recalculation.

---

## Part 2 — Using the Routines and Scripts

This section shows exactly how to invoke everything described in Part 1: the `phpp-tool` CLI and the two verification scripts. All commands assume an activated venv and a working directory at the project root.

```bash
cd /Users/smini/Documents/Coding/PHX_pyxl
source .venv/bin/activate
```

### `phpp-tool read` — extract a filled workbook to JSON

```bash
phpp-tool read Data/Example_IP.xlsx -o records/my_building.json
```

- `WORKBOOK` (positional) — path to the filled `.xlsx` to read.
- `-o, --output` — JSON output path (omit to print to stdout).
- `--phpp-version` — resolves `phpp-field-mapping/<version>.md` (default `EN_10_6_IP`). Use `EN_10_6_SI` for a genuinely SI-native single-shell workbook.
- `--field-map` — direct-path override, bypassing `--phpp-version` entirely.

Internally: `read_phpp()` → `BuildingRecord.from_reader_dict()` → `to_json()`, then stamps the output JSON with `_phpp_version` for the `write` command to cross-check later. Takes roughly 20 seconds for a full PHPP workbook.

### `phpp-tool write` — inject a JSON record into a blank template

```bash
phpp-tool write records/my_building.json Data/Empty_IP.xlsx -o output.xlsx
```

- `RECORD_FILE` (positional) — the JSON produced by `read`.
- `TEMPLATE` (positional) — the blank `.xlsx` to write into.
- `-o, --output` (required) — path for the written workbook.
- `--phpp-version` / `--field-map` — same as `read`. **Must match the version the record was read with** — `write` compares `--phpp-version` against the record's stamped `_phpp_version` and prints a warning (not a hard error) on mismatch.

Internally: `model_validate_json()` → `model_dump(exclude_none=True)` → `write_phpp()`. Takes roughly 30 seconds.

### `phpp-tool inspect-map` — audit field map coverage

```bash
phpp-tool inspect-map --phpp-version EN_10_6_IP
```

Prints every mapped worksheet with its field/section/config counts. Use this after editing a field map file to confirm the parser still finds everything (and that every config/items entry still has a valid type tag — `inspect-map` will raise `FieldMapError` immediately if one is missing).

### `scripts/roundtrip.py` — Parts 1 & 2 verification

```bash
python scripts/roundtrip.py Data/Example_IP.xlsx Data/Empty_IP.xlsx
python scripts/roundtrip.py Data/Empty_IP.xlsx Data/Empty_IP.xlsx
```

Arguments come in `source template` pairs — pass more pairs on the same command line to run several roundtrips in one invocation (results print a combined summary at the end). Output artifacts land in `records/roundtrip_<timestamp>/`.

What it runs, phase by phase:

| Phase | Command-line effect |
|-------|---------------------|
| 1–4 (Part 1, no Excel) | Read inputs, read all, write, verify every written cell against the writer's own write list |
| 5 (Part 2, optional) | Read the *source* live via xlwings+Excel and compare against openpyxl's cached values — skipped automatically if xlwings/Excel isn't available |

### `scripts/verify_excel.py` — Stage 3 full-fidelity comparison

```bash
python scripts/verify_excel.py Data/Example_IP.xlsx records/roundtrip_<timestamp>/Example_written.xlsx
```

Exactly two positional arguments: `source_file written_file`. Both files must already have been opened in Excel and saved manually — this refreshes cached formula values so openpyxl can read correct results without live Excel access. Output lands in `records/verify_excel_<timestamp>/`, including a machine-readable `verify_excel_summary.json`.

The full three-stage sequence in practice:

```bash
# Stages 1 & 2 — no manual step needed
python scripts/roundtrip.py Data/Example_IP.xlsx Data/Empty_IP.xlsx

# Manual step: open both Data/Example_IP.xlsx and the *_written.xlsx output in Excel, save, close

# Stage 3
python scripts/verify_excel.py Data/Example_IP.xlsx records/roundtrip_<timestamp>/Example_written.xlsx
```

### Running the test suite

```bash
pytest tests/ -v          # 88 tests, ~0.14s, no Excel needed
```

### Quick reference

| Task | Command |
|------|---------|
| Read a filled PHPP → JSON | `phpp-tool read Data/Example_IP.xlsx -o records/my_building.json` |
| Write JSON → blank PHPP | `phpp-tool write records/my_building.json Data/Empty_IP.xlsx -o output.xlsx` |
| Check field map coverage | `phpp-tool inspect-map` |
| Roundtrip verification (Parts 1–2) | `python scripts/roundtrip.py Data/Example_IP.xlsx Data/Empty_IP.xlsx` |
| Full-fidelity check (Stage 3, post-Excel-save) | `python scripts/verify_excel.py <source.xlsx> <written.xlsx>` |
| Unit tests | `pytest tests/ -v` |

---

## Part 3 — Routine-by-Routine Walkthrough

### `map_parser.py` — Field map parser

**Strategy role:** Converts the markdown field map into the structured dict that reader and writer iterate over.

| Step | What happens |
|------|-------------|
| 1 | `parse_field_map()` reads the markdown file and splits it at `##` headings (one per worksheet). |
| 2 | Each worksheet section is parsed for its worksheet key (e.g. `VERIFICATION`), sheet name, config block, label-anchored fields, and subsections. |
| 3 | Subsections (split at `###`) are parsed for header/entry locators, column fields, row fields, items, and appliance rows — capturing all six addressing strategies. |
| 4 | Returns a nested dict keyed by worksheet key, ready for reader and writer to iterate. |

### `locators.py` — Cell resolution (six strategies, openpyxl dual-load)

**Strategy role:** The bridge between the field map's abstract locator specs and actual cell values. Every reader and writer function calls into this module. Unlike PHX_xlwg (which uses xlwings range objects), all functions here operate on openpyxl `Worksheet` objects paired as `WsPair`.

#### Helpers

| Function | What it does |
|----------|-------------|
| `norm()` | Normalizes label text for comparison — NFKC unicode, NBSP→space, strip, casefold. Ensures labels match regardless of Excel's formatting quirks. |
| `col_to_idx()` | Converts column letters (`A`→1, `AA`→27) to 1-based numeric index for openpyxl cell access. |
| `field_col()` | Extracts the column letter from a field spec (handles both string and dict forms). |
| `_is_formula()` | Checks the `data_only=False` worksheet to determine whether a cell contains a formula (value starts with `=`). This is the core of formula detection. |
| `cell_value()` | Reads a single cell from the `data_only=True` worksheet. When `skip_formulas=True`, first checks `_is_formula()` against the `data_only=False` worksheet and returns `None` for formula cells. |
| `find_row_in_col()` | Searches for a text needle in a column using the `data_only=True` worksheet for label text matching. Iterates row by row from `start_from` to `max_row`. Supports substring and exact matching via `norm()`. |
| `parse_cell_ref()` | Splits `"AB123"` into `("AB", 123)`. |
| `is_header_row()` | Heuristic: a row is a header if all non-None values are strings. Used to skip header rows in block iteration. |
| `is_entry_row_header()` | Reads the entry row's column field values from the `data_only=True` worksheet and applies `is_header_row()` to determine if the entry locator points at a column header rather than the first data row. |

#### Strategy 1: Label-anchored — `resolve_label_anchored()`

Finds `locator_string` in `locator_col` using `find_row_in_col()` on `ws_vals` (the `data_only=True` worksheet), then reads the value at `input_col` in the found row (plus optional `row_offset`) via `cell_value()`. Applies `skip_formulas` filtering through the `WsPair`.

#### Strategy 2: Header + entry block — `resolve_block()`

The most complex resolver. Handles repeating data rows (e.g. window schedules, area entries).

| Step | What happens |
|------|-------------|
| 1 | Find the header row and entry row using `find_row_in_col()` on `ws_vals`. If the entry row is itself a column header (detected by `is_entry_row_header()`), start one row below. |
| 2 | Build a column index map from all `column_fields` for cell access. |
| 3 | Iterate from start row to `max_row`. For each row: read each column field from `ws_vals`, check `_is_formula()` against `ws_fmls` if `skip_formulas` is on. |
| 4 | Check for end markers (e.g. "Unhide additional rows"), detect sparse/empty rows (break after 3 consecutive), skip header rows, and collect data rows. |
| 5 | Return a list of row dicts, each with `_row` metadata for round-trip positioning. |

#### Strategy 3: Named range — `resolve_named_range()`

Looks up an Excel defined name via `wb_vals.defined_names[name]`, iterates its `destinations` to get `(title, coord)`, reads the cell value from `wb_vals[title][coord]`. If `skip_formulas` is on, also checks `wb_fmls[title][coord]` for formula strings. Returns `None` on missing names.

#### Strategy 4: Absolute address — `resolve_absolute()`

Reads `ws_vals[address].value` for a fixed cell reference like `"C11"`. If `skip_formulas` is on, checks `ws_fmls[address].value` for formula strings.

#### Strategy 5: Column + row-offset — `resolve_row_offset()`

Reads the cell at `col`, `anchor_row + row_offset`. Delegates to `cell_value()` with `skip_formulas` through the `WsPair`.

#### Strategy 6: Fixed result — `resolve_fixed()`

Reads a cell at a fixed `(row, col)`. Delegates to `cell_value()` with `skip_formulas`. Typically used for formula output cells (which are skipped when formula filtering is on).

### `reader.py` — openpyxl dual-load reader

**Strategy role:** Walks the field map, calls locator functions, and builds a nested dict of all input values from a filled PHPP workbook — without launching Excel.

#### Top level: `read_phpp()`

| Step | What happens |
|------|-------------|
| 1 | Load the workbook twice: `load_workbook(path, data_only=True)` → `wb_vals`, `load_workbook(path, data_only=False)` → `wb_fmls`. Both loads use default mode (not `read_only`, which causes hangs on large workbooks). |
| 2 | Parse the field map via `parse_field_map()`. |
| 3 | For each worksheet in the field map, find the matching sheet (preferring SI variants via `prefer_si_sheet()`). Pair the two worksheet objects into a `WsPair`. Call `_read_worksheet()`. |
| 4 | Close both workbooks. |
| 5 | Return the nested dict. |

#### `_read_worksheet()`

For a single sheet, reads three categories:

1. **Label-anchored fields** (`fields` key) — calls `_read_label_anchored_fields()`.
2. **Config values** (`config` key) — calls `_read_config()`.
3. **Sections** (`sections` key) — iterates sections, dispatching each via `_read_section()`.

All three pass `skip_formulas` and both workbook objects through.

#### `_read_section()` — dispatch

Examines which keys are present in the section spec (`header_locator`, `entry_locator`, `column_fields`, `row_fields`, `items`, `fields`, `appliance_rows`) and routes to the appropriate reader:

| Pattern | Reader | Locator strategies used |
|---------|--------|------------------------|
| header + entry + row_fields (no col_fields) | `_read_row_offset_section()` | Strategy 5 (column + row-offset) |
| col_fields + row_fields | `_read_column_row_section()` | Strategy 5 |
| header + entry + col_fields | `_read_block_section()` | Strategy 2 (header + entry block) |
| col_fields only (no header) | `_read_static_column_section()` | Strategy 5 |
| appliance_rows | `_read_appliance_section()` | Stub (metadata only) |
| items | `_read_items_section()` | Strategies 4, 3, 6 (absolute, named range, fixed) |
| fields only | `_read_label_anchored_fields()` | Strategy 1 (label-anchored) |
| header only | `_read_header_only()` | `find_row_in_col()` |

#### `_read_label_anchored_fields()`

Iterates the `fields` dict. For each field with a `locator_string`, calls `resolve_label_anchored()` (Strategy 1) with the `WsPair`. Collects results into a flat dict.

#### `_read_config()`

Iterates config key/value pairs. Classifies each via `classify_item()`: absolute addresses call `resolve_absolute()` (Strategy 4) with the `WsPair`, named ranges call `resolve_named_range()` (Strategy 3) with both `wb_vals` and `wb_fmls`, plain values pass through.

#### `_read_block_section()`

Delegates to `resolve_block()` (Strategy 2) with the `WsPair`. Returns a list of row dicts, one per data row in the repeating block.

#### `_read_row_offset_section()`

Finds the anchor row via `find_row_in_col()` on `ws_vals`, then reads each field at a row offset using `resolve_row_offset()` (Strategy 5) with the `WsPair`.

#### `_read_column_row_section()`

A grid pattern (e.g. DHW tanks). Iterates `column_fields` x `row_fields`, calling `resolve_row_offset()` (Strategy 5) with the `WsPair` for each intersection.

#### `_read_static_column_section()`

Iterates rows starting from `entry_row_start`, reading each `column_field` via `resolve_row_offset()` (Strategy 5) with the `WsPair`. Stops when an entire row is None.

#### `_read_items_section()`

Iterates an `items` dict. For dict items with `col`+`row`, calls `resolve_fixed()` (Strategy 6) with the `WsPair`. For string items, classifies as absolute address → `resolve_absolute()` (Strategy 4) or named range → `resolve_named_range()` (Strategy 3) with both workbook objects.

#### `_read_appliance_section()` (stub)

Returns metadata from the Electricity sheet's appliance row specs without reading live cell values. Does not require a `WsPair`.

#### `_read_header_only()`

Finds the header row position via `find_row_in_col()` on `ws_vals` and returns `{"_header_row": row}`.

### `writer.py` — openpyxl resolution + surgical XML persistence

**Strategy role:** Resolves cell addresses using a read-only `data_only=True` workbook (label text), collects all writes as `(sheet_name, col, row, value)` tuples, then hands them to `surgical_writer.apply_surgical_writes()` for persistence. No Excel required.

#### Top level: `write_phpp()`

| Step | What happens |
|------|-------------|
| 1 | Load the template with `data_only=True` → `wb_labels` (for label resolution only; never saved). |
| 2 | Parse the field map. For each worksheet key in the record, find the matching sheet and call `_write_worksheet()`, threading the plain sheet-name string through (no writable openpyxl object is needed) and collecting writes into a `pending` list. |
| 3 | Close `wb_labels`. |
| 4 | Call `surgical_writer.apply_surgical_writes(template_path, output_path, pending)` — this copies the template and patches only the `<sheetData>` of sheets with writes, as a ZIP/XML edit. See `surgical_writer.py` below. |
| 5 | Return the list of `(sheet_name, col, row, value)` writes for verification. |

#### `_write_cell()`

The gatekeeper: skips `None` and dict/list values, appends valid writes to the pending list as `(sheet_name, col, row, value)`.

#### `_write_worksheet()`

Mirrors `_read_worksheet()`: writes label-anchored fields first, then iterates sections via `_write_section()`.

#### `_write_section()` — dispatch

Same pattern detection as the reader's `_read_section()`, routing to:

- `_write_block()` — iterates row data, uses `_row` metadata or sequential offset to determine target rows, writes each column field.
- `_write_row_offset()` — finds the anchor row via `find_row_in_col()` on `ws_labels`, writes each field at its row offset.
- `_write_column_row()` — iterates the column x row grid, writes each intersection.
- `_write_static_column()` — iterates row data, writes each column field at the row's position.
- `_write_items()` — writes fixed-address, absolute-address, and named-range items.
- `_write_label_anchored()` — finds each label via `find_row_in_col()` on `ws_labels`, writes the value at the paired column.

#### `_write_named_range()`

Resolves the named range via `wb_labels.defined_names[name]`, uses `coordinate_from_string()` to extract `(col_letter, row_num)`, and appends to the pending list. Rejects `None`, dict, and list values.

### `surgical_writer.py` — ZIP/XML persistence (ported from PHX_Dev)

**Strategy role:** Persists the writer's `(sheet_name, col, row, value)` tuples without going through openpyxl's save cycle, so `<extLst>` extensions and `<headerFooter>` content survive untouched.

| Function | What it does |
|----------|-------------|
| `apply_surgical_writes()` | Public entry point. Groups the flat write list by sheet name, then calls `_apply_surgical()`. |
| `_apply_surgical()` | Copies the template if there are no writes. Otherwise builds a sheet-name → ZIP-path map, patches each affected sheet's XML, and rebuilds the ZIP. |
| `_build_sheet_map()` | Reads `xl/workbook.xml` for sheet names/rIds and `xl/_rels/workbook.xml.rels` for rId → file target, returning `{sheet_name: zip_path}`. |
| `_patch_sheet_xml()` | The core trick: finds the `<sheetData>...</sheetData>` substring by raw text search, parses the *whole* document with `lxml` only to build/modify `<row>`/`<c>` elements, then re-serializes *only* the `<sheetData>` element and splices it back into the original text at the same byte offsets. Everything before and after `<sheetData>` — including `<extLst>` and `<headerFooter>` — is never touched, re-parsed, or re-serialized. Restores `\r\n` line endings inside cached values after serialization, since lxml normalizes them to `\n` per the XML spec but Excel expects `\r\n` back. |
| `_set_cell_value()` | Chooses the correct XML encoding for a value: `<v>` for numbers, `<v t="b">` for booleans, `<is><t>` (inline string) for strings — and removes any existing `<f>` (formula) element, since a write always replaces a formula with a literal value. |
| `_sort_rows_and_cells()` | Excel requires strictly ascending row/cell order; re-sorts `<row>` elements by row number and `<c>` elements within each row after inserting new ones. |
| `_rebuild_zip()` | Streams the original ZIP into a new one, substituting only the modified sheet XML entries — every other archive member (styles, drawings, charts, VBA, custom XML) passes through as raw bytes, unmodified. |

### `models.py` — Pydantic validation

**Strategy role:** Provides a typed schema between the reader's raw dict and the JSON output. Validates and normalizes data without imposing rigid constraints on less-used worksheets.

| Step | What happens |
|------|-------------|
| 1 | Core worksheets (Verification, Overview, Climate, Ventilation, DHW, Windows) have explicit Pydantic models with typed fields. |
| 2 | Less-used worksheets use `dict[str, Any]` for flexibility. |
| 3 | `BuildingRecord.from_reader_dict()` maps worksheet keys to model classes, validates each, and assembles the top-level record. |
| 4 | `to_json()` serializes to JSON, excluding `None` values. |
| 5 | `model_validate_json()` (on the write path) deserializes and validates JSON back into a record. |

### `cli.py` — Click command-line interface

**Strategy role:** User-facing entry point that connects reader, models, and writer.

| Command | Pipeline |
|---------|----------|
| `phpp-tool read <filled.xlsx> -o record.json` | `read_phpp()` → `BuildingRecord.from_reader_dict()` → `to_json()` → write file |
| `phpp-tool write <record.json> <template.xlsx> -o output.xlsx` | `model_validate_json()` → `model_dump()` → `write_phpp()` |
| `phpp-tool inspect-map` | `parse_field_map()` → print worksheet/field/section counts |

### `scripts/roundtrip.py` — Two-part roundtrip verification

**Strategy role:** End-to-end test that proves data survives the full read → JSON → write cycle, with optional Excel-based cache validation.

#### Part 1 — Pure openpyxl (no Excel required)

| Phase | What happens |
|-------|-------------|
| 1 — Read inputs | `read_phpp(skip_formulas=True)` captures only input cells via the dual-load approach. Reports count of non-None values. |
| 2 — Read all | `read_phpp(skip_formulas=False)` captures everything. Compares against Phase 1 to report how many formula values were filtered (typically ~35%). |
| 3 — Write | `write_phpp()` writes inputs into the template using the dual-load writer. Returns the list of 13,102 cell writes. |
| 4 — Verify | Opens the written file with openpyxl (`data_only=False`) and checks every `(sheet, col, row, value)` tuple from the writer. Reports checked count, mismatches, and details. Both Example→Empty and Empty→Empty produce zero mismatches. |

#### Part 2 — Excel cache validation (optional, requires xlwings+Excel)

| Phase | What happens |
|-------|-------------|
| 5 — Excel read | Opens the *original* source workbook in a hidden Excel instance via xlwings. Reads all mapped fields with live formula recalculation. |
| 6 — Compare | Compares Excel's live values against openpyxl's cached values from Phase 2. If all values match, the cache is confirmed fresh — openpyxl was reading correct data. |

Part 2 is skipped gracefully if xlwings or Excel is not available. It validates openpyxl's cache accuracy, not the written file's formulas.

### `scripts/verify_excel.py` — Post-Excel full-fidelity comparison

**Strategy role:** The definitive test — compares *every* mapped field (inputs and formula results) between the original source and the written output, after both have been manually opened in Excel and saved.

**Prerequisite:** The user must manually open both files in Excel and save them before running this script. This refreshes all cached formula values so openpyxl can read correct results without programmatic Excel access.

| Phase | What happens |
|-------|-------------|
| 1 — Read both | Reads ALL cells from both files using `read_phpp(skip_formulas=False)`. Since both files were Excel-saved, every formula cell has a fresh cached value. |
| 2 — Deep compare | `_compare_deep()` recursively walks the nested dicts and lists from both reads, comparing every field. Reports mismatches with full dotted paths (e.g. `COMPONENTS.frames[0].u_value_bottom`). |

This is a separate script (not part of `roundtrip.py`) because it requires a manual step between the write and the verification — the user must open both files in Excel and save them. The three-step workflow is:

```
roundtrip.py  →  manual Excel open+save  →  verify_excel.py
```

If all inputs were transplanted correctly and Excel recalculated both files, every formula result should match.

---

## Part 4 — Usage, Roundtrip Philosophy, and Output Files

### Using the phpp_tool

The tool serves a single workflow: extract portable data from a filled PHPP workbook, and inject that data into a fresh PHPP template. This lets designers transfer building configurations between PHPP versions, share project data without sending 29 MB workbooks, archive inputs in a version-controllable text format, and programmatically generate or modify building records — all without requiring Excel to be installed.

#### Reading a filled workbook

```bash
phpp-tool read Data/Example_IP.xlsx -o records/my_building.json
```

No Excel installation required. The tool loads the workbook twice with openpyxl (once for cached values, once for formula detection), walks all 31 mapped worksheets, reads every input cell (skipping formulas), validates through Pydantic models, and writes a JSON file. The process takes roughly 20 seconds for a full PHPP workbook.

The resulting JSON is organized by worksheet key. Each worksheet contains some combination of scalar fields, config values, and sections. A section may be a flat dict of key-value pairs, a list of row dicts (for repeating blocks like window schedules), or a grid of column x row intersections.

#### Writing into a blank template

```bash
phpp-tool write records/my_building.json Data/Empty_IP.xlsx -o output.xlsx
```

No Excel installation required. The tool opens the template read-only with `data_only=True` for label resolution, resolves all cell addresses using the same locator strategies as the reader, collects writes, then persists them via `surgical_writer.py`'s ZIP/XML patch rather than an openpyxl save. This takes roughly 12 seconds (down from ~30s under the old openpyxl-save approach, since patching only the `<sheetData>` region of affected sheets is cheaper than re-serializing the whole workbook).

The output file is a valid .xlsx that contains all the designer's input values in the correct cells, with `<extLst>` extensions (Data Validation, etc.) and `<headerFooter>` content preserved byte-for-byte from the template — verified across all 83 sheets of a full roundtrip write.

#### Inspecting the field map

```bash
phpp-tool inspect-map
```

Lists every mapped worksheet with counts of fields, sections, and config items. Useful for verifying field map coverage after edits.

### Comparison with PHX_xlwg

| Concern | PHX_xlwg (xlwings) | PHX_pyxl (openpyxl) |
|---------|-------------------|---------------------|
| Excel requirement | Yes (read + write) | No |
| Formula values | Live recalculation | Cached (from last Excel save) |
| Formula detection | `cell.formula` property via xlwings | Dual-load: `data_only=False` returns formula string |
| Write strategy | xlwings address resolution + surgical XML persistence | openpyxl (`data_only=True`) address resolution + surgical XML persistence |
| File integrity | Equivalent — both persist via the same `surgical_writer.py` ZIP/XML patch, preserving `<extLst>`/`<headerFooter>` | Equivalent (see left) |
| Read performance | ~21s (Excel launch + AppleScript overhead) | ~20s (pure Python XML parsing) |
| Write performance | ~48s (Excel launch dominates; persistence itself is fast) | ~12s (openpyxl resolution + surgical patch) |
| Test suite speed | ~9s (requires Excel) | ~0.16s (no Excel) |
| Headless/CI operation | No | Yes |

### The roundtrip test: philosophy

The roundtrip test answers a single question: **does every input value survive the full read → JSON → write cycle unchanged?**

This is not a unit test of individual functions — those are covered by the 88 pytest cases. The roundtrip test is an integration test of the entire pipeline against real PHPP workbooks. It treats the tool as a black box and verifies the output against the input at the cell level.

#### Why cell-by-cell verification matters

The tool touches 13,102 cells across 31 worksheets. A mismatch in any one of them could mean a wrong U-value, a missing ventilation rate, or a misplaced area entry. Aggregate statistics (like "99.9% match") would hide single-cell errors that could be significant in a Passive House certification. The test therefore checks every cell individually and reports exact addresses for any mismatch.

#### Why two reads (inputs vs all)

The roundtrip reads the source workbook twice:

1. **Inputs only** (`skip_formulas=True`) — this is what gets written to JSON and into the template. It captures only designer-entered values.
2. **All cells** (`skip_formulas=False`) — this captures everything including formula results. Comparing the two reveals the formula filter's effect.

This dual read serves two purposes. First, it quantifies the formula filter: for the Example workbook, 7,625 of 22,044 non-None values (34.6%) are formula results that would overwrite formulas if written back. Second, it documents which worksheet fields are formulas versus inputs — the Verification sheet, for example, is entirely formula-driven:

| Field | All cells | Inputs only |
|-------|-----------|-------------|
| `phi_building_category_type` | `"21-Non-res building: School half-days (< 7 h)"` | `None` |
| `phi_certification_type` | `"10-Passive house"` | `None` |
| `setpoint_winter` | `20.0` | `None` |
| ... | (13 fields with values) | (all None — every field is a formula) |

This means the Verification sheet's display values are computed from inputs on other sheets. The tool correctly skips them.

#### Why the writer returns its write list

The writer's `write_phpp()` function returns the list of `(sheet_name, col, row, value)` tuples it collected during the label-resolution phase and persisted to the output workbook. This is not an incidental convenience — it is the verification contract. The roundtrip test opens the written file with openpyxl and checks every tuple against the actual cell contents. If the writer says it wrote `42.0` to `Ventilation SI!J35`, the test confirms that cell `J35` on sheet `Ventilation SI` contains `42.0`.

#### Why three-stage verification

The verification pipeline has three stages, each building on the last:

**Stage 1 — `roundtrip.py` Part 1** (pure openpyxl) proves that every input value written to the output file persists correctly — the write-then-read-back loop is bit-accurate. But it cannot verify that openpyxl's cached formula values were correct in the first place, nor can it verify that formula results in the written file are correct (since openpyxl-written formula cells have no cached values until Excel recalculates them).

**Stage 2 — `roundtrip.py` Part 2** (xlwings+Excel) addresses the cache freshness question by reading the *original* source workbook in Excel, which recalculates all formulas live. Comparing Excel's live values against openpyxl's cached values from Stage 1 confirms that the cache is fresh — openpyxl was reading the correct data. If any cached value is stale, Part 2 reports the discrepancy. However, Part 2 reads only the original source, not the written output, because openpyxl-written files cannot be reliably opened by Excel via AppleScript.

**Stage 3 — `verify_excel.py`** (post-Excel, pure openpyxl) closes the loop. After the user has manually opened *both* the original source and the written output in Excel and saved them, all cached formula values in both files are fresh. `verify_excel.py` then reads both files with openpyxl (`data_only=True`, `skip_formulas=False`) and compares *every* mapped field — inputs and formula results alike. If all inputs were written correctly and Excel recalculated both files, formula results should match between source and output. This is the definitive full-fidelity test, confirming that the written workbook produces the same calculation results as the original — without requiring programmatic Excel access.

```bash
# Stages 1 & 2:
python scripts/roundtrip.py Data/Example_IP.xlsx Data/Empty_IP.xlsx

# Manual step: open both files in Excel, save, close.

# Stage 3:
python scripts/verify_excel.py Data/Example_IP.xlsx records/.../Example_written.xlsx
```

#### The two standard test pairs

| Test | Source | Template | What it proves |
|------|--------|----------|---------------|
| Example → Empty | A filled PHPP with real building data | A blank PHPP template | Input values from a real project survive the cycle |
| Empty → Empty | A blank PHPP template | The same blank template | Default/structural values survive; no spurious data is introduced |

Both produce zero mismatches across 13,102 verified cells.

### Output files

#### CLI output: JSON building record

Produced by `phpp-tool read`. Structure:

```
{
  "VERIFICATION": {                    <- worksheet key
    "phi_building_category_type": null, <- formula cell (skipped)
    "setpoint_winter": null,            <- formula cell (skipped)
    ...
  },
  "COMPONENTS": {
    "glazings": [                       <- repeating block (list of row dicts)
      {
        "_row": 115,                    <- source row in the Excel sheet
        "id": "1187gl03",
        "description": "EAGON - EAGON SUPER VIG (5/0,25 Vac/:5 Vac.)",
        "g_value": 0.48,
        "u_value": 0.51
      },
      ...                               <- 170 rows for glazings alone
    ],
    "frames": [ ... ],
    "ventilators": [ ... ]
  },
  "CLIMATE": {
    "named_ranges": {                   <- named range values
      "country": "US-United States of America",
      "region": "New York",
      "data_set": "New York/JFK"
    },
    ...
  },
  "DHW": {
    "tanks": {                          <- column x row grid
      "tank_1": {
        "tank_type": null,
        "standby_losses": null,
        ...
      },
      ...
    }
  },
  ...                                   <- 17 worksheet keys total
}
```

Key characteristics:

- **`_row` metadata** — Block rows carry their source row number so the writer can place them back at the correct position, even if the template has a different row layout.
- **`null` values** — Formula cells and empty input cells both appear as `null`. The writer skips `null` values, so formulas are preserved in the output workbook.
- **No formula text** — The JSON never contains formula strings like `=SUM(...)`. Only resolved input values appear.
- **`_config` sections** — Store worksheet-level settings (active column selections, variant references) that configure how the writer interprets the data.

#### CLI output: written workbook (.xlsx)

Produced by `phpp-tool write`. This is a copy of the blank template with input values injected at the addresses the writer resolved. Formulas remain intact as they were in the template, and `<extLst>`/`<headerFooter>` content is preserved byte-for-byte since persistence goes through `surgical_writer.py`'s ZIP/XML patch rather than an openpyxl save. The file can be opened in Excel manually with no degraded features.

#### Roundtrip test artifacts

**`roundtrip.py`** — saved to `records/roundtrip_<timestamp>/`:

| File | Contents |
|------|----------|
| `<name>_inputs.json` | The input-only read (`skip_formulas=True`). This is the data that travels through the pipeline — the same output `phpp-tool read` would produce. 14,419 non-None values for the Example workbook. |
| `<name>_all.json` | The full read (`skip_formulas=False`). Includes formula results alongside inputs. 22,044 non-None values for the Example workbook. Comparing against `_inputs.json` shows exactly which cells are formulas. |
| `<name>_written.xlsx` | The written workbook. Template with input values injected. This is what `phpp-tool write` would produce. The roundtrip test verifies every cell in this file against the writer's reported write list. |
| `<name>_orig_excel.json` | (Part 2 only) Excel live-recalculated values from the original source. Comparing against `_all.json` validates openpyxl's cached formula values. |

**`verify_excel.py`** — saved to `records/verify_excel_<timestamp>/`:

| File | Contents |
|------|----------|
| `<source>_all.json` | Full read of the Excel-saved original source (`skip_formulas=False`). All cached values are fresh from the manual Excel save. |
| `<written>_all.json` | Full read of the Excel-saved written file (`skip_formulas=False`). Formula results are now cached from Excel's recalculation. |
| `verify_excel_summary.json` | Machine-readable summary: field counts, match/mismatch totals, timing. |

The console reports include:

- **`roundtrip.py`**: worksheet and value counts for both reads, formula filter statistics, number of cell writes, cell-by-cell verification result, and cache validation result (Part 2)
- **`verify_excel.py`**: worksheet and value counts for both files, full-fidelity comparison result (fields checked, matched, mismatched, with dotted-path details for each mismatch)

---

## Part 5 — Concerns, Features, and Limitations

Lessons from building both PHX_pyxl and its sibling PHX_xlwg against the same field map, with openpyxl and xlwings respectively.

### `phpp-field-mapping.md`

**Features**

- Single shared dictionary — both PHX_pyxl and PHX_xlwg read the exact same markdown file and get identical locator behavior, so the field map only needs maintaining once.
- Human-readable and git-diffable — field map edits review like code changes, not like an opaque binary spreadsheet-to-spreadsheet mapping.
- Six addressing strategies (label-anchored, header+entry block, named ranges, absolute address, column+row-offset, fixed result) cover PHPP's inconsistent per-sheet layout without needing sheet-specific code in the reader/writer.
- `phpp-tool inspect-map` gives instant field/section/config coverage auditing after any field map edit.

**Concerns / Limitations**

- ~~**SI/IP unit mismatch**~~ — **Fixed 2026-07-01.** The field map is now versioned: `phpp-field-mapping/EN_10_6_IP.md` and `EN_10_6_SI.md`, selected via `--phpp-version` (default `EN_10_6_IP`). `prefer_si_sheet()`'s runtime "try `<Name> SI` first" guessing is deleted entirely — each version declares its own correct `sheet_name` and locator strings directly. This turned out to matter more than a label mismatch: in an IP-shell workbook the `<Name> SI` tabs are formula mirrors of the base tab's real inputs (see the new item below), so the old default silently lost data, not just risked a wrong unit label. Verified against genuinely distinct files: `Data/Example_IP.xlsx` (IP-shell, dual-tab) and `Data/Example_SI.xlsx` (SI-native, single-shell, no `SI`-suffixed tabs at all) — full details and evidence in `phpp-concerns-and-examples.md` #1.
- **Depends on German internal Excel defined names** (e.g. `Werte_Klima_Region`) — fragile if a PHPP version localizes differently or renames internal ranges.
- **Assumes a stable PHPP layout** — absolute-address and fixed-result strategies hardcode row/column numbers. A PHPP template revision that inserts or removes rows silently breaks these locators with no built-in detection or warning.
- **Stale relative to test workbooks** — the label `"DHW circulation pipes or, for heat interface units, forward and return flows"` is defined in the map but not found in either test PHPP workbook, suggesting the map has drifted from the PHPP versions actually in use.
- ~~**Conflates inputs and outputs**~~ — **Fixed 2026-07-01.** Label-anchored entries can now carry an optional `` `key` (input) ``/`` `key` (output) `` tag, cross-checked against actual cell formula status at read time (warns on drift) and enforced at write time (`writer.py` refuses to write to `(output)`-tagged fields). A one-time migration auto-tagged all 20 label-anchored entries in both `EN_10_6_IP.md`/`EN_10_6_SI.md` against the blank templates. Surprising result while building this: the concern's own original two examples (`phi_building_category_type`, `setpoint_winter`) no longer reproduce — both are literal on the base `Verification` tab, and were only "formula-driven" because of the `<Name> SI` mirror-tab bug below (now fixed), not because they're genuinely output fields. All 20 tags ended up `(input)`, none `(output)` — see `phpp-concerns-and-examples.md` #5.

**Specific data-quality defects (verified 2026-07-01):**

- **Malformed `phi_certification_class` row** (Verification, field-map line 21) — unescaped `|` characters inside the label and options text make this a non-standard markdown table row. It doesn't currently break anything: `map_parser.py`'s `_parse_label_row()` uses a state machine specifically hardened for this row (see its docstring), and the parsed `locator_string` — `"Class | Primary energy method"` — matches the real label text in `Verification SI!T13`. Still, it's fragile: any future change to the label-row parser that doesn't account for embedded pipes would silently break this field.
- **`energy_unit: KHW` typo** (SolarDHW config, field-map line 680) — should be `KWH`, as used everywhere else including the structurally identical PV config block two sections later. Confirmed harmless: `energy_unit`/`footprint_unit` are never read by any Python module — they're descriptive metadata only, not consumed by the reader, writer, or models.
- ~~**Climate `ud_block` header locator is broken**~~ — **Fixed 2026-07-01, and turned out to be a code bug, not bad data.** `PH-Tools/PHX`'s own shape file has the identical "swapped-looking" `header_locator` shape for this section, which was the tell. The real bugs: `start_row: 67` wasn't recognized as an alias for `entry_row_start`/`entry_start_row`, so it was silently ignored; and `_read_section()`'s dispatch order checked `has_items` before checking for `column_fields` anchored by an explicit start row, misrouting to `_read_items_section()`. Fixed both (alias added, dispatch reordered) in `reader.py`/`writer.py` in both PHX_pyxl and PHX_xlwg. `CLIMATE.ud_block` now resolves real monthly climate data — see `phpp-concerns-and-examples.md` #8.
- **Duplicate target cells — one confirmed intentional, one still a genuine bug.** `psi_g_left`/`psi_g_right`/`psi_g_bottom`/`psi_g_top` (Windows → frames, lines 290–293) all mapping to column `IR` is **documented as intentional** in `PHX_Dev/CLAUDE.md` — the original prototype's planning doc — since PHPP treats the glazing-edge (spacer) thermal bridge as one uniform value per window, unlike the installation thermal bridge (`psi_i`), which genuinely varies by side. `duct_assign_1`–`8` (Ventilation, lines 477–484) mapping sequentially to columns Q–X, then `duct_assign_9`/`_10` (lines 485–486) both jumping to `Z` (skipping `Y`) has no such documented rationale and remains a genuine copy-paste-style defect — every duct row silently loses whatever distinct value actually lives in column `Y`.

**Structural concerns affecting efficiency/correctness (verified 2026-07-01):**

- ~~**Config value type inferred from string shape, not declared**~~ — **Fixed 2026-07-01.** Every config/items bullet now requires an explicit `(literal)`/`(address)`/`(named_range)` tag; `map_parser.py` raises `FieldMapError` at parse time if one is missing or invalid, rather than falling back to shape-based guessing. `classify_item()`, `_is_cell_ref()`, and `_is_named_range()` are deleted entirely from `locators.py` — the parsed tag is looked up directly at read/write time. A one-time migration script tagged all 170 existing entries from their then-current classification; the 2 known-wrong `footprint_unit` entries were hand-corrected to `(literal)`. `SOLAR_DHW.footprint_unit`/`SOLAR_PV.footprint_unit` now read back as `"M2"` — see `phpp-concerns-and-examples.md` #10.
- **`UVALUES` and `EASY_PH` are dead worksheet entries.** Their `sheet_name` values (`"U-Values"`, `"easyPH"`) don't match any real sheet in either test workbook — the actual sheets are `"U-values SI"`/`"R-Values"` (different capitalization and naming scheme entirely), and there's no `easyPH` tab at all in PHPP 10.6. Both are silently skipped on every read and write (an INFO-level log line easy to miss in normal output). The commonly quoted "31 mapped worksheets" figure is optimistic — at least 2 of them currently do nothing.
- ~~**The `options` enum metadata is mostly inert.**~~ — **Fixed 2026-07-01.** `reader.py` now checks every resolved label-anchored value's leading code against its field's documented `options` dict and logs a warning on mismatch. Fires correctly on the genuine pre-existing drift in `phi_building_category_type` (resolves to code `21`, not among its documented `1`/`2`/`11`/`12`) and stays silent for fields whose resolved code matches (e.g. `phi_certification_type`) — see `phpp-concerns-and-examples.md` #12.
- ~~**Minor: the field map is re-parsed from scratch on every call.**~~ — **Fixed 2026-07-01.** `parse_field_map()` now caches by resolved path + mtime; a second call against the same on-disk file returns the identical (by `is`) dict object rather than re-parsing — see `phpp-concerns-and-examples.md` #13.

**Additional structural concerns (verified 2026-07-01, second pass):**

- ~~**`ADDNL_VENT` is silently dropped in its entirety under the default `skip_formulas=True` mode**~~ — **Fixed 2026-07-01. Was the most severe defect found in the first verification pass.** `resolve_block()`'s sparse-row heuristic discarded a row if it had no string value and very few non-`None` fields — meant to detect genuinely blank template rows. Under `skip_formulas=True`, formula-driven fields like `display_name` got nulled out *before* this check ran, tripping the same heuristic on real, populated room rows; two levels up, bare `if result:` truthiness checks then dropped the whole worksheet key silently. **The fix:** sparseness is now decided against raw, unfiltered values first, with `skip_formulas` applied only to what's actually returned; also added a diagnostic warning when a mapped worksheet resolves to no data at all. `read_phpp('Data/Example_IP.xlsx', ...)['ADDNL_VENT']['rooms']` now returns real populated rows — see `phpp-concerns-and-examples.md` #14.
- ~~**`entry_row_start`, when present, silently overrides the discovered entry-locator row with no cross-check.**~~ — **Fixed 2026-07-01.** Both `resolve_block()` and `_read_column_row_section()` (the `tanks` section actually dispatches through the latter, not the former — a good reminder that similarly-shaped sections can take genuinely different code paths) still use `entry_row_start` when present, but now also discover the label's actual row and log a warning when the two disagree. The `tanks` section's own drift is real and still present in the field map: `entry_row_start: 191` vs. the label `"Storage type 1"` actually sitting at row 189 — that's a field-map data question (which number is correct), left unresolved by design, since the fix's job was making the drift *visible*, not deciding which value wins. See `phpp-concerns-and-examples.md` #15.

**New findings and versioning architecture (2026-07-01, third pass — see `phpp-concerns-and-examples.md` #24-25 for full evidence):**

- ~~**`<Name> SI` tabs in IP-shell workbooks are formula mirrors of the base tab, not independent data**~~ — **Fixed 2026-07-01.** Sampling formula-vs-value ratios across all 28 worksheet pairs in `Example_IP.xlsx` showed a minority of worksheets where the SI tab has noticeably more formula cells than the base tab — exactly where genuine input cells live on the base tab and get mirrored onto the SI tab via `=IF(...)`-style passthrough formulas for unit-converted display. Since `skip_formulas` treats any formula cell as "not a real input," the old `prefer_si_sheet()` default silently discarded real designer input workbook-wide, not just in `ADDNL_VENT`. Fixed by deleting `prefer_si_sheet()` and moving to per-version explicit sheet names (see the SI/IP concern above).
- ~~**`Climate.ud_block`'s `summer_delta_t_unit` had a bogus `DELTA-C` "column" value**~~ — **Fixed 2026-07-01.** Every sibling row in that column-fields table maps to a real 1-3 letter column; this one had the literal string `DELTA-C` — a stray unit annotation in the wrong table cell. Harmless on PHX_pyxl (`col_to_idx("DELTA-C")` silently produces a nonsensical but non-crashing index) but crashed PHX_xlwg's live AppleScript call outright when discovered during the Stage D port. Removed from all four field map copies (`EN_10_6_SI.md`/`EN_10_6_IP.md` × PHX_pyxl/PHX_xlwg) — it never represented a real column mapping and nothing consumed the key.
- **Field map is now versioned.** `phpp-field-mapping/EN_10_6_IP.md` and `EN_10_6_SI.md` replace the single `phpp-field-mapping.md`, selected via `--phpp-version` (default `EN_10_6_IP`). Adopted from the architectural pattern in `PH-Tools/PHX`'s `phpp_localization/` directory (one shape file per language/version/unit-variant) — though nothing was copied from that GPL-3.0 project; both files here are independently authored and verified against this project's own `Example_IP.xlsx`/`Example_SI.xlsx`.

### openpyxl

**Features**

- Pure Python — no Excel installation, no OS automation (AppleScript/COM), runs headless and in CI.
- Fast — roughly 20s for a full-workbook read, and the 88-test suite runs in ~0.14s with no Excel dependency.
- The dual-load pattern (`data_only=True`/`False`) gets both label text and formula-vs-input classification without ever launching Excel.
- Deterministic — read results depend only on the file's saved XML, never on live application or session state.

**Concerns / Limitations**

- **Cached, not live, formula values** — reads reflect whatever Excel last saved, not necessarily current data. Confirming freshness requires the optional xlwings+Excel comparison in roundtrip Part 2.
- ~~**Save cycle drops content via two distinct mechanisms**~~ — **Fixed 2026-07-01.** This concern applied to *openpyxl's own save cycle*: (1) `parse_extensions()` unconditionally discards any of 8 GUID-tagged `<extLst>` extension types it recognizes — Conditional Formatting, Data Validation, Sparkline Group, Slicer List, Protected Range, Ignored Error, Web Extension, Timeline Ref (`openpyxl/xml/constants.py:EXT_TYPES`) — of which only **Data Validation** was confirmed to actually fire on `Example.xlsx`/`Empty.xlsx`. (2) `header_footer.py`'s `_split_string()` regex fails to match PHPP's header/footer format codes and silently blanks all three sections. The writer no longer calls openpyxl's `save()` at all — `surgical_writer.py` persists writes via a ZIP/XML patch instead, touching only the `<sheetData>` region of affected sheets. Verified byte-for-byte: `<extLst>` and `<headerFooter>` regions are identical between template and written output across all 83 sheets of a full 13,102-cell roundtrip write. The only remaining caveat is unrelated: written files still can't be reopened by Excel via AppleScript automation on macOS 26 (a separate OS/Excel-version compatibility issue, not something the writer controls).
- **Written files can't be reopened via AppleScript automation** — data validation errors cause a hang, so written files can only be reopened manually in the Excel GUI.
- **No recalculation** — full-fidelity verification of formula results (`verify_excel.py`) requires a manual step (open both files in Excel, save) that pure openpyxl cannot automate away.

### xlwings

**Features**

- Drives a live Excel instance, so formulas always recalculate — PHX_pyxl's roundtrip Part 2 uses this to confirm openpyxl's cached values are still fresh.
- No file-format degradation on read — Excel serializes its own file, so reading via xlwings preserves every extension openpyxl would otherwise strip.

**Concerns / Limitations**

- **Hard Excel dependency** — unusable in headless or CI environments; PHX_pyxl deliberately avoids this as its primary backend for exactly that reason, using xlwings only as an optional validation step.
- **macOS 26 broke every native AppleScript save path** for large workbooks (`wb.save()` errors, `close(saving=yes)` hangs indefinitely) — this regression is the reason PHX_pyxl exists as an openpyxl-only alternative to its sibling project.
- **Slower per operation** once Excel launch and AppleScript RPC overhead are counted (~21s read / ~48s write in PHX_xlwg vs. ~20s / ~30s here) — the optional Part 2 cache check adds this overhead on top of the openpyxl path it validates.
- **Operational fragility** — alert dialogs, transient AppleScript error -50 on rapid open/close, and Excel processes left running after abnormal exit all need explicit handling that a pure-Python approach avoids entirely.
- **New this session, PHX_xlwg-specific:** xlwings' Mac AppleScript backend can silently drop a row's worth of data from a batch `.value`/`.formula` read when the range crosses certain hidden-row boundaries — found and fixed in PHX_xlwg's `locators.py` (see that project's own strategy doc and `phpp-concerns-and-examples.md` #26 for full detail). Not directly exercised by PHX_pyxl's own optional Part 2 xlwings check, but a useful data point on the class of correctness risk that comes with batch-reading via Excel automation, as opposed to openpyxl's direct XML parsing.
