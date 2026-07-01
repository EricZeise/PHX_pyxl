# PHX_pyxl: Strategy and Routine Reference

## Part 1 — Strategy Summary

### Goal

Read designer-entered input data from a filled PHPP workbook (Passive House Planning Package, ~29 MB, 83 sheets), store it as portable JSON, and write that JSON back into a blank PHPP template — preserving every input value while leaving all formulas untouched. Do this entirely with openpyxl, requiring no Excel installation.

### Why openpyxl (and why not xlwings)

The sibling project PHX_xlw uses xlwings to drive a live Excel instance for reading and writing. This works well on machines where Excel is installed, but introduces hard dependencies: Excel must be present, AppleScript (macOS) or COM (Windows) must be functional, and the process is bound to a GUI application's lifecycle. macOS 26 introduced an AppleScript bug where Excel save operations hang indefinitely on large workbooks, forcing PHX_xlw into a hybrid architecture.

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

### The writer's dual-load approach

The writer also loads the template workbook twice, but for a different reason:

| Load | `data_only` | Purpose |
|------|-------------|---------|
| `wb_labels` | `True` | Resolve locator lookups — label searches, named ranges, entry-row detection — using cached text values |
| `wb_out` | `False` | Apply cell writes while preserving all formulas in untouched cells |

The writable load (`data_only=False`) preserves formula strings in cells the writer doesn't touch. The label-resolution load (`data_only=True`) provides readable text for the same locator strategies the reader uses.

### The field map

A single markdown file (`phpp-field-mapping.md`) serves as the dictionary of every mapped cell across 31 worksheets. It defines six addressing strategies for locating cells, since PHPP's layout varies from sheet to sheet:

1. **Label-anchored** — find a text label in one column, read/write the value in a paired column at an optional row offset.
2. **Header + entry block** — find a header row and entry row, then iterate repeating data rows reading each column field.
3. **Named ranges** — resolve an Excel defined name (often German, e.g. `Werte_Klima_Region`) to its cell.
4. **Absolute address** — a fixed cell reference like `C11`.
5. **Column + row-offset** — locate an anchor row, then read at a column with a fixed row offset.
6. **Fixed result rows/cols** — a fixed row and column, typically for result cells.

### Known limitation: openpyxl save degradation

openpyxl's load→save cycle strips some Excel extensions: data validation rules, custom headers/footers, and certain XML namespace extensions that PHPP relies on. The written file is a valid .xlsx and can be opened manually in Excel, but:

- It cannot be reopened by Excel via AppleScript (the data validation errors cause a hang)
- Some PHPP drop-down menus and validation constraints may be missing
- The file is best treated as a data artifact, not a production PHPP workbook

### Verification

The roundtrip test confirms data fidelity end to end. The writer returns its full list of `(sheet_name, col, row, value)` writes. The test script opens the written file with openpyxl and verifies every cell matches — 13,102 cells, zero mismatches. An optional second part uses xlwings+Excel to validate that openpyxl's cached formula values match Excel's live recalculation.

---

## Part 2 — Routine-by-Routine Walkthrough

### `map_parser.py` — Field map parser

**Strategy role:** Converts the markdown field map into the structured dict that reader and writer iterate over.

| Step | What happens |
|------|-------------|
| 1 | `parse_field_map()` reads the markdown file and splits it at `##` headings (one per worksheet). |
| 2 | Each worksheet section is parsed for its worksheet key (e.g. `VERIFICATION`), sheet name, config block, label-anchored fields, and subsections. |
| 3 | Subsections (split at `###`) are parsed for header/entry locators, column fields, row fields, items, and appliance rows — capturing all six addressing strategies. |
| 4 | Returns a nested dict keyed by worksheet key, ready for reader and writer to iterate. |

### `locators.py` — Cell resolution (six strategies, openpyxl dual-load)

**Strategy role:** The bridge between the field map's abstract locator specs and actual cell values. Every reader and writer function calls into this module. Unlike PHX_xlw (which uses xlwings range objects), all functions here operate on openpyxl `Worksheet` objects paired as `WsPair`.

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

### `writer.py` — Pure openpyxl writer

**Strategy role:** Resolves cell addresses using the `data_only=True` workbook (label text), collects all writes as tuples, applies them to the `data_only=False` workbook (preserving formulas), and saves. No Excel required.

#### Top level: `write_phpp()`

| Step | What happens |
|------|-------------|
| 1 | Copy the template file to the output path. |
| 2 | Load the template with `data_only=True` → `wb_labels` (for label resolution). Load the output copy with `data_only=False` → `wb_out` (for writing, preserving formulas). |
| 3 | Parse the field map. For each worksheet key in the record, find the matching sheet and call `_write_worksheet()`, collecting writes into a `pending` list. |
| 4 | Close `wb_labels`. |
| 5 | Apply all pending writes to `wb_out` by setting `cell(row, column).value` for each `(sheet_name, col, row, value)` tuple. Save and close `wb_out`. |
| 6 | Return the list of `(sheet_name, col, row, value)` writes for verification. |

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

## Part 3 — Usage, Roundtrip Philosophy, and Output Files

### Using the phpp_tool

The tool serves a single workflow: extract portable data from a filled PHPP workbook, and inject that data into a fresh PHPP template. This lets designers transfer building configurations between PHPP versions, share project data without sending 29 MB workbooks, archive inputs in a version-controllable text format, and programmatically generate or modify building records — all without requiring Excel to be installed.

#### Reading a filled workbook

```bash
phpp-tool read Data/Example.xlsx -o records/my_building.json
```

No Excel installation required. The tool loads the workbook twice with openpyxl (once for cached values, once for formula detection), walks all 31 mapped worksheets, reads every input cell (skipping formulas), validates through Pydantic models, and writes a JSON file. The process takes roughly 20 seconds for a full PHPP workbook.

The resulting JSON is organized by worksheet key. Each worksheet contains some combination of scalar fields, config values, and sections. A section may be a flat dict of key-value pairs, a list of row dicts (for repeating blocks like window schedules), or a grid of column x row intersections.

#### Writing into a blank template

```bash
phpp-tool write records/my_building.json Data/Empty.xlsx -o output.xlsx
```

No Excel installation required. The tool copies the template, opens the template with `data_only=True` for label resolution, opens the copy with `data_only=False` for writing, resolves all cell addresses using the same locator strategies as the reader, collects writes, applies them, and saves. This takes roughly 30 seconds.

The output file is a valid .xlsx that contains all the designer's input values in the correct cells. However, because openpyxl's save strips some Excel extensions (data validation rules, custom headers/footers), the file should be treated as a data artifact. It can be opened manually in Excel, but some PHPP features may be degraded.

#### Inspecting the field map

```bash
phpp-tool inspect-map
```

Lists every mapped worksheet with counts of fields, sections, and config items. Useful for verifying field map coverage after edits.

### Comparison with PHX_xlw

| Concern | PHX_xlw (xlwings) | PHX_pyxl (openpyxl) |
|---------|-------------------|---------------------|
| Excel requirement | Yes (read + write) | No |
| Formula values | Live recalculation | Cached (from last Excel save) |
| Formula detection | `cell.formula` property via xlwings | Dual-load: `data_only=False` returns formula string |
| Write strategy | xlwings address resolution + openpyxl persistence | Dual-load: `data_only=True` for labels, `data_only=False` for writing |
| File integrity | Better (xlwings resolves, openpyxl only persists values) | Degraded (openpyxl load→save strips extensions) |
| Read performance | ~21s (Excel launch + AppleScript overhead) | ~20s (pure Python XML parsing) |
| Write performance | ~48s (Excel launch + openpyxl persistence) | ~30s (two openpyxl loads + save) |
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
python scripts/roundtrip.py Data/Example.xlsx Data/Empty.xlsx

# Manual step: open both files in Excel, save, close.

# Stage 3:
python scripts/verify_excel.py Data/Example.xlsx records/.../Example_written.xlsx
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

Produced by `phpp-tool write`. This is a copy of the blank template with input values injected at the addresses the writer resolved. Formulas remain intact as they were in the template. The file can be opened in Excel manually, though some PHPP data validation features may be degraded due to openpyxl's load→save cycle stripping certain Excel extensions.

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
