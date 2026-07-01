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
phpp-field-mapping.md   (locator dictionary — where each field lives)
        ↓
   map_parser.py        (parse markdown → structured dict)
        ↓
   locators.py          (6 addressing strategies, openpyxl dual-load)
        ↓
 ┌──────┴──────┐
 reader.py    writer.py  (openpyxl resolution + surgical XML persistence)
 └──────┬──────┘
     models.py           (pydantic validation)
        ↓
     cli.py              (Click CLI)
```

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
├── pyproject.toml               ← deps: openpyxl, lxml, click, pydantic (NO xlwings)
├── phpp-field-mapping.md        ← the locator dictionary (31 worksheets)
├── src/
│   ├── phpp_tool/
│   │   ├── __init__.py
│   │   ├── cli.py               ← Click CLI: read / write / inspect-map
│   │   ├── map_parser.py        ← Parse phpp-field-mapping.md → structured dict
│   │   ├── locators.py          ← 6 addressing strategies (openpyxl dual-load)
│   │   ├── reader.py            ← openpyxl-based reader
│   │   ├── writer.py            ← Resolves addresses via openpyxl, collects writes
│   │   ├── surgical_writer.py   ← Persists writes via ZIP/XML patch (lxml), preserving extLst/headerFooter
│   │   └── models.py            ← Pydantic models for building record JSON
│   └── compare_json/            ← Standalone JSON diff tool
├── scripts/
│   ├── roundtrip.py             ← Two-part roundtrip test (Parts 1 & 2)
│   └── verify_excel.py          ← Post-Excel full-fidelity comparison
├── tests/
└── records/                     ← Output directory for JSON building records
```

---

## Commands

```bash
# Install in dev mode
pip install -e ".[dev]"

# Run tests (no Excel needed — 88 tests, <1 second)
pytest tests/ -v

# Read a PHPP into JSON
phpp-tool read path/to/filled_PHPP.xlsx -o records/my_building.json

# Write a record into a blank PHPP
phpp-tool write records/my_building.json path/to/blank_PHPP.xlsx -o output.xlsx

# Roundtrip test (Part 1: openpyxl only; Part 2: xlwings+Excel if available)
python scripts/roundtrip.py Data/Example.xlsx Data/Empty.xlsx

# Post-Excel verification (after manually opening both files in Excel and saving)
python scripts/verify_excel.py Data/Example.xlsx records/.../Example_written.xlsx
```

---

## Constraints

- **Never read or embed PHPP formulas** — only designer-entered input values.
- **The field map is the single source of truth** for cell locations.
- **No Excel required** for reading, writing, or Part 1 verification.
- **Writes preserve file integrity** — persistence goes through `surgical_writer.py`'s ZIP/XML patch, not openpyxl's save, so `<extLst>` extensions (Data Validation, etc.) and `<headerFooter>` content survive the round trip. Verified byte-for-byte across all 83 sheets of a full roundtrip write.
- **Cached formula values** — without Excel, formula results reflect the last Excel save, not a live recalculation. Part 2 of the roundtrip test uses xlwings+Excel (optional) to verify recalculation integrity.
