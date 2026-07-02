"""Parse phpp-field-mapping.md into structured locator definitions.

The field map markdown uses six addressing strategies (label-anchored,
header+entry blocks, named ranges, absolute addresses, row-offset blocks,
and fixed rows/cols). This module parses all of them into a nested dict
keyed by worksheet_key -> sections -> fields.
"""

from __future__ import annotations

import logging
import re
from pathlib import Path
from typing import Any

logger = logging.getLogger(__name__)


class FieldMapError(Exception):
    """Raised when the field map markdown is malformed."""


_parse_cache: dict[str, tuple[float, dict[str, Any]]] = {}


def parse_field_map(path: str | Path) -> dict[str, Any]:
    """Parse the PHPP field mapping markdown into a structured dict.

    Returns a dict keyed by worksheet key (e.g. ``"VERIFICATION"``).
    Each value has ``sheet_name``, ``config``, ``fields``,
    and ``sections``.

    Cached by resolved path + mtime -- a single process (e.g.
    scripts/roundtrip.py, which calls this once per read/write) only
    re-parses the ~900-line markdown file once per distinct on-disk state.
    Callers must not mutate the returned dict; it's shared across calls.
    """
    resolved = str(Path(path).resolve())
    mtime = Path(resolved).stat().st_mtime
    cached = _parse_cache.get(resolved)
    if cached is not None and cached[0] == mtime:
        return cached[1]

    text = Path(resolved).read_text(encoding="utf-8")
    result: dict[str, Any] = {}
    for heading, body in _split_on_heading(text, level=2):
        ws = _parse_worksheet(heading, body)
        if ws is not None:
            result[ws["worksheet_key"]] = ws

    _parse_cache[resolved] = (mtime, result)
    return result


# ---------------------------------------------------------------------------
# Structure splitting
# ---------------------------------------------------------------------------

def _split_on_heading(text: str, level: int) -> list[tuple[str, str]]:
    """Split *text* at markdown headings of *level* (2 = ``##``, 3 = ``###``).

    Returns ``[(heading_text, body), ...]``.
    """
    prefix = "#" * level + " "
    parts = re.split(rf"^{re.escape(prefix)}", text, flags=re.MULTILINE)
    result: list[tuple[str, str]] = []
    for part in parts[1:]:
        first_line, _, rest = part.partition("\n")
        result.append((first_line.strip(), rest))
    return result


# ---------------------------------------------------------------------------
# Worksheet parsing
# ---------------------------------------------------------------------------

def _parse_worksheet(heading: str, body: str) -> dict[str, Any] | None:
    key_match = re.search(r"Worksheet key:\s*`(\w+)`", body)
    if not key_match:
        return None

    worksheet_key = key_match.group(1)

    first_h3 = body.find("\n### ")
    top_content = body[:first_h3] if first_h3 != -1 else body

    config, config_kind = _extract_config(top_content)
    fields = _extract_label_anchored_fields(top_content)

    sections: dict[str, Any] = {}
    for sec_name, sec_body in _split_on_heading(body, level=3):
        sections[sec_name] = _parse_section(sec_body)

    return {
        "worksheet_key": worksheet_key,
        "sheet_name": heading,
        "config": config,
        "config_kind": config_kind,
        "fields": fields,
        "sections": sections,
    }


# ---------------------------------------------------------------------------
# Section parsing
# ---------------------------------------------------------------------------

def _parse_section(body: str) -> dict[str, Any]:
    """Parse an H3 subsection body into a structured dict."""
    locators = _extract_locators(body)
    items, items_kind = _extract_items(body)
    tables = _extract_tables(body)

    column_fields: dict[str, Any] = {}
    row_fields: dict[str, Any] = {}
    label_fields: dict[str, Any] = {}
    appliance_rows: dict[str, Any] = {}

    for table in tables:
        kind = table["kind"]
        if kind == "label_anchored":
            label_fields.update(table["fields"])
        elif kind == "column":
            column_fields.update(table["fields"])
        elif kind == "row":
            row_fields.update(table["fields"])
        elif kind == "row_offset":
            row_fields.update(table["fields"])
        elif kind == "appliance":
            appliance_rows.update(table["fields"])

    section: dict[str, Any] = {}
    if locators:
        section.update(locators)
    if items:
        section["items"] = items
        section["items_kind"] = items_kind
    if column_fields:
        section["column_fields"] = column_fields
    if row_fields:
        section["row_fields"] = row_fields
    if label_fields:
        section["fields"] = label_fields
    if appliance_rows:
        section["appliance_rows"] = appliance_rows
    return section


# ---------------------------------------------------------------------------
# Table extraction
# ---------------------------------------------------------------------------

_TABLE_LINE = re.compile(r"^\|.+\|$")
_SEPARATOR_LINE = re.compile(r"^\|[\s\-:|]+\|$")


def _extract_tables(text: str) -> list[dict[str, Any]]:
    """Find all markdown tables in *text* and parse each."""
    lines = text.split("\n")
    tables: list[dict[str, Any]] = []
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        if _TABLE_LINE.match(line) and i + 1 < len(lines):
            next_line = lines[i + 1].strip()
            if _SEPARATOR_LINE.match(next_line):
                headers = _split_table_line(line)
                num_cols = len(headers)
                norm_h = [h.lower().replace("`", "").strip() for h in headers]
                is_label = any(
                    "locator_string" in h or h == "label" for h in norm_h
                )
                data_rows: list[list[str]] = []
                j = i + 2
                while j < len(lines):
                    row_line = lines[j].strip()
                    if not _TABLE_LINE.match(row_line):
                        break
                    ec = None if is_label else num_cols
                    data_rows.append(_split_table_line(row_line, ec))
                    j += 1
                tables.append(_classify_table(headers, data_rows))
                i = j
                continue
        i += 1
    return tables


def _split_table_line(
    line: str, expected_cols: int | None = None,
) -> list[str]:
    """Split a markdown table row on ``|``."""
    parts = line.split("|")
    if parts and not parts[0].strip():
        parts = parts[1:]
    if parts and not parts[-1].strip():
        parts = parts[:-1]
    cells = [p.strip() for p in parts]

    if expected_cols is not None and len(cells) > expected_cols:
        overflow = " | ".join(cells[expected_cols - 1:])
        cells = cells[: expected_cols - 1] + [overflow]
    return cells


def _classify_table(
    headers: list[str], rows: list[list[str]]
) -> dict[str, Any]:
    """Classify a table by its headers and parse rows accordingly."""
    norm = [h.lower().replace("`", "").strip() for h in headers]

    if any("locator_string" in h or h == "label" for h in norm):
        return _parse_label_table(rows)

    if "appliance" in norm or "data row" in norm:
        return _parse_appliance_table(headers, rows)

    has_column = "column" in norm
    has_row_offset = "row offset" in norm
    has_row = "row" in norm and not has_row_offset

    if has_row_offset:
        return _parse_row_offset_table(headers, rows)
    if has_row and has_column:
        col_fields = _parse_column_table(headers, rows)
        row_fields = _parse_row_table(headers, rows)
        combined: dict[str, Any] = {}
        combined.update(col_fields.get("fields", {}))
        combined.update(row_fields.get("fields", {}))
        return {"kind": "column", "fields": combined}
    if has_row:
        return _parse_row_table(headers, rows)
    if has_column:
        return _parse_column_table(headers, rows)

    return {"kind": "unknown", "headers": headers, "rows": rows}


# --- Label-anchored table (Strategy 1) ---

def _parse_label_table(rows: list[list[str]]) -> dict[str, Any]:
    fields: dict[str, Any] = {}
    for row_cells in rows:
        parsed = _parse_label_row(row_cells)
        if parsed:
            key = parsed.pop("field_key")
            fields[key] = parsed
    return {"kind": "label_anchored", "fields": fields}


def _parse_label_row(cells: list[str]) -> dict[str, Any] | None:
    """Parse a single label-anchored table row using a state machine.

    Robust to extra columns (like the malformed ``phi_certification_class``
    row) and pipe characters inside option values.
    """
    field_key: str | None = None
    field_io: str | None = None
    label_parts: list[str] = []
    locator_col: str | None = None
    input_col: str | None = None
    row_offset: int = 0
    unit: str | None = None
    options_parts: list[str] = []

    state = "find_key"
    for cell in cells:
        stripped = cell.strip()

        if state == "find_key":
            m = re.match(r"^`(.+?)`\s*(?:\((input|output)\))?$", stripped)
            if m:
                field_key = m.group(1)
                field_io = m.group(2)
                state = "label"
            continue

        if state == "label":
            if re.match(r"^[A-Z]{1,3}$", stripped):
                locator_col = stripped
                state = "input_col"
            elif stripped:
                label_parts.append(stripped)
            continue

        if state == "input_col":
            if re.match(r"^[A-Z]{1,3}$", stripped):
                input_col = stripped
                state = "row_offset"
            continue

        if state == "row_offset":
            if re.match(r"^\d+$", stripped):
                row_offset = int(stripped)
            state = "unit"
            continue

        if state == "unit":
            unit = stripped if stripped else None
            state = "options"
            continue

        if state == "options":
            if stripped:
                options_parts.append(cell)

    if field_key is None:
        return None

    options_text = " | ".join(options_parts).strip()
    return {
        "field_key": field_key,
        "locator_string": " | ".join(label_parts) if label_parts else None,
        "locator_col": locator_col,
        "input_col": input_col,
        "row_offset": row_offset,
        "unit": unit,
        "options": _parse_options(options_text) if options_text else None,
        "io": field_io,
    }


# --- Column field table ---

def _parse_column_table(
    headers: list[str], rows: list[list[str]]
) -> dict[str, Any]:
    norm = [h.lower().strip() for h in headers]
    col_idx = _find_header(norm, "column")
    unit_idx = _find_header(norm, "unit")
    field_idx = _find_header(norm, "field")

    fields: dict[str, Any] = {}
    for cells in rows:
        name = _cell(cells, field_idx).strip("`").strip()
        if not name:
            continue
        entry: dict[str, Any] = {"column": _cell(cells, col_idx)}
        if unit_idx is not None:
            u = _cell(cells, unit_idx)
            entry["unit"] = u if u else None
        fields[name] = entry
    return {"kind": "column", "fields": fields}


# --- Row field table ---

def _parse_row_table(
    headers: list[str], rows: list[list[str]]
) -> dict[str, Any]:
    norm = [h.lower().strip() for h in headers]
    row_idx = _find_header(norm, "row")
    unit_idx = _find_header(norm, "unit")
    field_idx = _find_header(norm, "field")

    fields: dict[str, Any] = {}
    for cells in rows:
        name = _cell(cells, field_idx).strip("`").strip()
        if not name:
            continue
        row_val = _cell(cells, row_idx)
        entry: dict[str, Any] = {"row": _coerce(row_val)}
        if unit_idx is not None:
            u = _cell(cells, unit_idx)
            entry["unit"] = u if u else None
        fields[name] = entry
    return {"kind": "row", "fields": fields}


# --- Row-offset field table ---

def _parse_row_offset_table(
    headers: list[str], rows: list[list[str]]
) -> dict[str, Any]:
    norm = [h.lower().strip() for h in headers]
    offset_idx = next(
        i for i, h in enumerate(norm) if "row offset" in h
    )
    unit_idx = _find_header(norm, "unit")
    field_idx = _find_header(norm, "field")

    fields: dict[str, Any] = {}
    for cells in rows:
        name = _cell(cells, field_idx).strip("`").strip()
        if not name:
            continue
        offset_val = _cell(cells, offset_idx)
        entry: dict[str, Any] = {"row_offset": _coerce(offset_val)}
        if unit_idx is not None:
            u = _cell(cells, unit_idx)
            entry["unit"] = u if u else None
        fields[name] = entry
    return {"kind": "row_offset", "fields": fields}


# --- Appliance rows table (Electricity input_rows) ---

def _parse_appliance_table(
    headers: list[str], rows: list[list[str]]
) -> dict[str, Any]:
    norm = [h.lower().strip() for h in headers]
    app_idx = _find_header(norm, "appliance")
    data_idx = _find_header(norm, "data row")
    sel_idx = _find_header(norm, "selection row")
    opt_idx = _find_header(norm, "options")

    fields: dict[str, Any] = {}
    for cells in rows:
        name = _cell(cells, app_idx).strip("`").strip()
        if not name:
            continue
        entry: dict[str, Any] = {}
        dr = _cell(cells, data_idx)
        if dr:
            entry["data_row"] = _coerce(dr)
        sr = _cell(cells, sel_idx) if sel_idx is not None else ""
        if sr:
            entry["selection_row"] = _coerce(sr)
        opts = _cell(cells, opt_idx) if opt_idx is not None else ""
        if opts:
            entry["options"] = _parse_options(opts)
        fields[name] = entry
    return {"kind": "appliance", "fields": fields}


# ---------------------------------------------------------------------------
# Bullet / item extraction
# ---------------------------------------------------------------------------

_BULLET_ITEM = re.compile(
    r"^-\s+`([^`]+)`\s*(?:\(([a-z_]+)\))?\s*:\s*(.+)$", re.MULTILINE
)

_LOCATOR_PATTERN = re.compile(
    r"^-\s+(Header|Entry)\s+locator:\s*col\s+`([^`]+)`"
    r"\s*,\s*string\s+`\"([^\"]*)\"`",
    re.MULTILINE,
)

_COL_ROW_PATTERN = re.compile(
    r"^-\s+`([^`]+)`\s*:\s*col\s+`([^`]+)`"
    r"\s*,\s*row\s+`(\d+)`(?:\s*\((\w+)\))?",
    re.MULTILINE,
)

_VALID_KINDS = {"literal", "address", "named_range"}
_CELL_REF_RE = re.compile(r"^[A-Z]{1,3}\d+$")
_NAMED_RANGE_RE = re.compile(r"^[A-Z][A-Za-z0-9]*(_[A-Za-z0-9]+)+$")


def _extract_locators(text: str) -> dict[str, Any]:
    """Extract header_locator and entry_locator from bullet patterns."""
    result: dict[str, Any] = {}
    for m in _LOCATOR_PATTERN.finditer(text):
        kind = m.group(1).lower()
        result[f"{kind}_locator"] = {
            "col": m.group(2),
            "string": m.group(3),
        }
    return result


def _extract_items(text: str) -> tuple[dict[str, Any], dict[str, str]]:
    """Extract bullet-list key: value items (excluding locator lines).

    Every plain bullet item must declare its kind explicitly --
    ``literal``, ``address``, or ``named_range`` -- as
    `` - `key` (kind): `value` ``. Raises FieldMapError if a bullet is
    missing a valid kind tag, rather than silently guessing from the
    value's shape.

    Returns (items, kinds): items maps key -> coerced value (unchanged
    from before), kinds maps key -> its declared kind, for callers that
    need to resolve addresses/named ranges without re-inferring the type.
    """
    items: dict[str, Any] = {}
    kinds: dict[str, str] = {}

    for m in _COL_ROW_PATTERN.finditer(text):
        entry: dict[str, Any] = {"col": m.group(2), "row": int(m.group(3))}
        if m.group(4):
            entry["unit"] = m.group(4)
        items[m.group(1)] = entry

    col_row_spans = {(m.start(), m.end())
                     for m in _COL_ROW_PATTERN.finditer(text)}
    locator_spans = {(m.start(), m.end())
                     for m in _LOCATOR_PATTERN.finditer(text)}
    skip_spans = col_row_spans | locator_spans

    for m in _BULLET_ITEM.finditer(text):
        if any(s <= m.start() < e for s, e in skip_spans):
            continue
        key = m.group(1)
        tag = m.group(2)
        raw = m.group(3).strip().strip("`").strip('"').strip()
        if tag not in _VALID_KINDS:
            raise FieldMapError(
                f"Field map entry `{key}` is missing a required type tag "
                f"-- one of {sorted(_VALID_KINDS)} -- in line: {m.group(0)!r}"
            )
        value = _coerce(raw)
        items[key] = value
        kinds[key] = tag

        if isinstance(value, str):
            looks_address = bool(_CELL_REF_RE.match(value))
            looks_named_range = bool(_NAMED_RANGE_RE.match(value))
            if tag == "literal" and (looks_address or looks_named_range):
                logger.warning(
                    "Field map entry `%s` is tagged (literal) but value "
                    "%r looks like a %s -- check this wasn't a mistake",
                    key, value,
                    "cell address" if looks_address else "named range",
                )
            elif tag == "address" and not looks_address:
                logger.warning(
                    "Field map entry `%s` is tagged (address) but value "
                    "%r doesn't look like a cell reference", key, value)
            elif tag == "named_range" and not looks_named_range:
                logger.warning(
                    "Field map entry `%s` is tagged (named_range) but "
                    "value %r doesn't look like a named range", key, value)

    return items, kinds


def _extract_config(text: str) -> tuple[dict[str, Any], dict[str, str]]:
    """Extract configuration items from the top-level worksheet body."""
    return _extract_items(text)


# ---------------------------------------------------------------------------
# Label-anchored field extraction (top-level tables)
# ---------------------------------------------------------------------------

def _extract_label_anchored_fields(text: str) -> dict[str, Any]:
    """Parse label-anchored tables found in the top-level worksheet body."""
    fields: dict[str, Any] = {}
    for table in _extract_tables(text):
        if table["kind"] == "label_anchored":
            fields.update(table["fields"])
    return fields


# ---------------------------------------------------------------------------
# Options parsing
# ---------------------------------------------------------------------------

def _parse_options(text: str) -> dict[str, str]:
    """Parse an options string like ```1`: val1; `2`: val2`` into a dict."""
    opts: dict[str, str] = {}
    if not text:
        return opts
    for part in text.split(";"):
        part = part.strip()
        if not part:
            continue
        m = re.match(r"`?(\w+)`?\s*:\s*(.+)", part)
        if m:
            opts[m.group(1)] = m.group(2).strip()
    return opts


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _find_header(norm_headers: list[str], keyword: str) -> int | None:
    for i, h in enumerate(norm_headers):
        if keyword in h:
            return i
    return None


def _cell(cells: list[str], idx: int | None) -> str:
    if idx is None or idx >= len(cells):
        return ""
    return cells[idx].strip()


def _coerce(value: str) -> int | float | str:
    """Coerce a string to int or float if possible."""
    if not value:
        return value
    try:
        return int(value)
    except ValueError:
        pass
    try:
        return float(value)
    except ValueError:
        pass
    return value
