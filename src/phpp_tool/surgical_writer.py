"""Surgical ZIP/XML cell-value writer for PHPP workbooks.

Ported from PHX_Dev's writer.py. Instead of openpyxl's load->modify->save
cycle (which drops <extLst> extensions like Data Validation, and mangles
<headerFooter> content it can't parse), this module edits the .xlsx file
as a ZIP archive: it patches only the <sheetData> region of the specific
sheet XML files that contain cells being written, leaving every other byte
of the file -- including <extLst> and <headerFooter> -- untouched.
"""

from __future__ import annotations

import logging
import re
import shutil
import zipfile
from io import BytesIO
from pathlib import Path
from typing import Any

from lxml import etree

logger = logging.getLogger(__name__)

CellWrite = tuple[str, int, Any]  # (col_letter, row_num, value)
SheetWrites = dict[str, list[CellWrite]]  # sheet_name -> list of writes
SkippedRef = tuple[str, str, int]  # (sheet_name, col_letter, row_num)

SHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def apply_surgical_writes(
    template_path: str | Path,
    output_path: str | Path,
    writes: list[tuple[str, str, int, Any]],
) -> set[SkippedRef]:
    """Apply (sheet_name, col, row, value) writes via surgical XML patch.

    Copies the template byte-for-byte, then edits only the <sheetData>
    of sheets that have writes -- everything else (extLst, headerFooter,
    drawings, charts, VBA) passes through untouched.

    Returns the set of (sheet_name, col, row) writes that were skipped
    because the target cell already held a formula in the template --
    overwriting it would silently delete that formula (and, if it was a
    shared-formula master, orphan every dependent cell in its range,
    which is what makes Excel flag the file for repair on open). Callers
    should drop these refs from their own returned write list so
    verification doesn't expect a value that was deliberately not
    written.
    """
    template_path = Path(template_path)
    output_path = Path(output_path)

    by_sheet: SheetWrites = {}
    for sheet_name, col, row, value in writes:
        by_sheet.setdefault(sheet_name, []).append((col, row, value))

    return _apply_surgical(template_path, output_path, by_sheet)


def _apply_surgical(
    template_path: Path, output_path: Path, writes: SheetWrites
) -> set[SkippedRef]:
    """Copy template and apply cell-level XML edits to specific sheets."""
    if not writes:
        shutil.copy2(template_path, output_path)
        return set()

    sheet_to_zip_path = _build_sheet_map(template_path)

    modified_zip_paths: dict[str, bytes] = {}
    skipped: set[SkippedRef] = set()
    with zipfile.ZipFile(template_path, "r") as zf_in:
        for target_sheet, cell_writes in writes.items():
            zip_path = sheet_to_zip_path.get(target_sheet)
            if zip_path is None:
                continue
            original_xml = zf_in.read(zip_path)
            modified_xml, sheet_skipped = _patch_sheet_xml(
                original_xml, target_sheet, cell_writes)
            modified_zip_paths[zip_path] = modified_xml
            skipped.update(sheet_skipped)

    _rebuild_zip(template_path, output_path, modified_zip_paths)
    return skipped


def _build_sheet_map(xlsx_path: Path) -> dict[str, str]:
    """Build a mapping from sheet display name -> ZIP entry path.

    Reads workbook.xml for sheet names/rIds, then workbook.xml.rels
    for rId -> file target.
    """
    with zipfile.ZipFile(xlsx_path, "r") as zf:
        wb_xml = zf.read("xl/workbook.xml")
        wb_root = etree.fromstring(wb_xml)

        rels_xml = zf.read("xl/_rels/workbook.xml.rels")
        rels_root = etree.fromstring(rels_xml)

    rid_to_target: dict[str, str] = {}
    for rel in rels_root:
        rid_to_target[rel.get("Id", "")] = rel.get("Target", "")

    sheet_map: dict[str, str] = {}
    for sheet_el in wb_root.iter(f"{{{SHEET_NS}}}sheet"):
        name = sheet_el.get("name", "")
        rid = sheet_el.get(f"{{{REL_NS}}}id", "")
        target = rid_to_target.get(rid, "")
        if target:
            if not target.startswith("/"):
                target = "xl/" + target
            else:
                target = target.lstrip("/")
            sheet_map[name] = target

    return sheet_map


def _cell_ref(col: str, row: int) -> str:
    """Build a cell reference string like 'C11'."""
    return f"{col.upper()}{row}"


def _parse_ref(ref: str) -> tuple[int, int]:
    """Parse 'AB123' -> (col_number, row_number)."""
    m = re.match(r"^([A-Z]+)(\d+)$", ref)
    if not m:
        return (0, 0)
    col_str, row_str = m.group(1), m.group(2)
    col_num = 0
    for ch in col_str:
        col_num = col_num * 26 + (ord(ch) - ord("A") + 1)
    return col_num, int(row_str)


def _patch_sheet_xml(
    original_xml: bytes, sheet_name: str, cell_writes: list[CellWrite]
) -> tuple[bytes, list[SkippedRef]]:
    """Modify cells, preserving bytes outside <sheetData>.

    Strategy: parse the full document with lxml to modify cells, but only
    re-serialize the <sheetData> section. The original bytes outside that
    section stay untouched, which preserves namespace attribute order on
    <worksheet>, <extLst> extensions, <headerFooter> content, and all other
    formatting. Inside <sheetData>, \\r\\n in cached formula values is
    restored after serialization (lxml normalizes \\r\\n -> \\n per the XML
    spec, but Excel expects its original \\r\\n back).

    A write targeting a cell that already holds a formula in the template
    is skipped rather than applied: blindly overwriting it would delete
    that formula, and if the cell happens to be the master of a shared
    formula group (``<f t="shared" ref="...">``), every dependent cell in
    that range loses its definition, which is exactly what makes Excel
    flag the written file for repair on open. This is a defense-in-depth
    check -- the field map is supposed to only ever target designer-input
    cells -- so a skip here means an upstream locator bug reached past
    its intended range, not that this is expected behavior.
    """
    original_text = original_xml.decode("UTF-8")

    sd_open_pos = original_text.find("<sheetData")
    if sd_open_pos == -1:
        return original_xml, []

    sd_close_tag = "</sheetData>"
    sd_close_pos = original_text.find(sd_close_tag)
    if sd_close_pos == -1:
        return original_xml, []
    sd_end_pos = sd_close_pos + len(sd_close_tag)

    original_sd_text = original_text[sd_open_pos:sd_end_pos]
    has_crlf = "\r\n" in original_sd_text

    root = etree.fromstring(original_xml)
    ns = f"{{{SHEET_NS}}}"

    sheet_data = root.find(f"{ns}sheetData")
    if sheet_data is None:
        return original_xml, []

    row_elements: dict[int, etree._Element] = {}
    for row_el in sheet_data.findall(f"{ns}row"):
        r = row_el.get("r")
        if r and r.isdigit():
            row_elements[int(r)] = row_el

    skipped: list[SkippedRef] = []

    for col_letter, row_num, value in cell_writes:
        ref = _cell_ref(col_letter, row_num)

        row_el = row_elements.get(row_num)
        if row_el is None:
            row_el = etree.SubElement(sheet_data, f"{ns}row")
            row_el.set("r", str(row_num))
            row_elements[row_num] = row_el

        cell_el = None
        for c in row_el.findall(f"{ns}c"):
            if c.get("r") == ref:
                cell_el = c
                break

        if cell_el is not None and cell_el.find(f"{ns}f") is not None:
            logger.warning(
                "Refusing to overwrite formula cell %r!%s with %r -- "
                "the field map likely resolved a row range past its "
                "intended block, skipping this write",
                sheet_name, ref, value,
            )
            skipped.append((sheet_name, col_letter, row_num))
            continue

        if cell_el is None:
            cell_el = etree.SubElement(row_el, f"{ns}c")
            cell_el.set("r", ref)

        _set_cell_value(cell_el, value, ns)

    _sort_rows_and_cells(sheet_data, ns)

    new_sd = etree.tostring(sheet_data, encoding="unicode")
    if has_crlf:
        new_sd = new_sd.replace("\n", "\r\n")

    patched = (original_text[:sd_open_pos] + new_sd +
               original_text[sd_end_pos:]).encode("UTF-8")
    return patched, skipped


def _set_cell_value(cell_el: etree._Element, value: Any, ns: str) -> None:
    """Set a cell element's value, choosing the correct XML type encoding.

    - Numbers (int/float): <v>value</v>, no type attribute
    - Booleans: <v>1/0</v>, t="b"
    - Strings: <is><t>value</t></is>, t="inlineStr"
    """
    for child in list(cell_el):
        local = etree.QName(child.tag).localname
        if local in ("v", "is"):
            cell_el.remove(child)

    f_el = cell_el.find(f"{ns}f")

    if isinstance(value, bool):
        cell_el.set("t", "b")
        if f_el is not None:
            cell_el.remove(f_el)
        v_el = etree.SubElement(cell_el, f"{ns}v")
        v_el.text = "1" if value else "0"

    elif isinstance(value, (int, float)):
        if "t" in cell_el.attrib:
            del cell_el.attrib["t"]
        if f_el is not None:
            cell_el.remove(f_el)
        v_el = etree.SubElement(cell_el, f"{ns}v")
        v_el.text = repr(value) if isinstance(value, float) else str(value)

    elif isinstance(value, str):
        cell_el.set("t", "inlineStr")
        if f_el is not None:
            cell_el.remove(f_el)
        is_el = etree.SubElement(cell_el, f"{ns}is")
        t_el = etree.SubElement(is_el, f"{ns}t")
        t_el.text = value

    else:
        if "t" in cell_el.attrib:
            del cell_el.attrib["t"]
        if f_el is not None:
            cell_el.remove(f_el)
        v_el = etree.SubElement(cell_el, f"{ns}v")
        v_el.text = str(value)


def _sort_rows_and_cells(sheet_data: etree._Element, ns: str) -> None:
    """Sort <row> elements by row number and <c> elements within each row.

    Excel requires rows and cells in ascending order.
    """
    for row_el in sheet_data.findall(f"{ns}row"):
        cells = row_el.findall(f"{ns}c")
        for c in cells:
            row_el.remove(c)
        for c in sorted(cells, key=lambda c: _parse_ref(c.get("r", "A0"))):
            row_el.append(c)
    sheet_data[:] = sorted(sheet_data, key=lambda r: int(r.get("r", "0")))


def _rebuild_zip(
    source_path: Path,
    output_path: Path,
    modified: dict[str, bytes],
) -> None:
    """Rebuild the ZIP, replacing only the entries in *modified*.

    Unmodified entries are copied as-is to preserve fidelity.
    """
    buf = BytesIO()
    with zipfile.ZipFile(source_path, "r") as zf_in:
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf_out:
            for item in zf_in.infolist():
                if item.filename in modified:
                    zf_out.writestr(item, modified[item.filename])
                else:
                    zf_out.writestr(item, zf_in.read(item.filename))

    output_path.write_bytes(buf.getvalue())
