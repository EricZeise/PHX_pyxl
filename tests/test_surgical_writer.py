"""Tests for the surgical ZIP/XML cell-value writer."""

from __future__ import annotations

import zipfile
from pathlib import Path

from openpyxl import Workbook

from phpp_tool.surgical_writer import _patch_sheet_xml, apply_surgical_writes

_SHEET_XML = (
    b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    b'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    b"<sheetData>"
    b'<row r="1">'
    # A1 is the master of a shared-formula group spanning A1:A3.
    b'<c r="A1" t="str"><f t="shared" ref="A1:A3" si="0">B1</f><v>1</v></c>'
    b'<c r="C1"/>'
    b"</row>"
    b'<row r="2">'
    # A2 is a dependent member: references si=0 but carries no formula text.
    b'<c r="A2"><f t="shared" si="0"/><v>2</v></c>'
    b"</row>"
    b"</sheetData>"
    b"</worksheet>"
)


class TestFormulaProtection:
    """A field-map bug can resolve writes past a block's true end and land
    on cells that are still formulas in the target template. Overwriting
    the master of a shared-formula group orphans every dependent cell in
    its range, which is exactly what makes Excel flag the file for repair
    on open -- so these writes must be skipped, not applied.
    """

    def test_skips_write_to_shared_formula_master(self):
        patched, skipped = _patch_sheet_xml(_SHEET_XML, "Sheet1", [("A", 1, 99)])

        assert skipped == [("Sheet1", "A", 1)]
        assert b'<f t="shared" ref="A1:A3" si="0">B1</f>' in patched
        assert b">99<" not in patched

    def test_skips_write_to_shared_formula_dependent(self):
        patched, skipped = _patch_sheet_xml(_SHEET_XML, "Sheet1", [("A", 2, 99)])

        assert skipped == [("Sheet1", "A", 2)]
        assert b'<f t="shared" si="0"/>' in patched
        assert b">99<" not in patched

    def test_allows_write_to_non_formula_cell(self):
        patched, skipped = _patch_sheet_xml(_SHEET_XML, "Sheet1", [("C", 1, 42)])

        assert skipped == []
        assert b"<v>42</v>" in patched

    def test_mixed_batch_skips_only_the_formula_cell(self):
        patched, skipped = _patch_sheet_xml(
            _SHEET_XML, "Sheet1", [("A", 1, 99), ("C", 1, 42)])

        assert skipped == [("Sheet1", "A", 1)]
        assert b"<v>42</v>" in patched
        assert b'<f t="shared" ref="A1:A3" si="0">B1</f>' in patched


class TestApplySurgicalWrites:
    def test_end_to_end_skips_formula_cell_and_preserves_it(self, tmp_path: Path):
        template = tmp_path / "template.xlsx"
        output = tmp_path / "output.xlsx"

        wb = Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        ws["C1"] = None
        wb.save(template)

        # Inject a shared-formula master at A1 directly into the saved
        # template's XML, mirroring the real PHPP structure (openpyxl has
        # no API for authoring shared formulas).
        with zipfile.ZipFile(template, "r") as zf:
            names = zf.namelist()
            sheet_path = next(n for n in names if n.startswith("xl/worksheets/sheet"))
            contents = {n: zf.read(n) for n in names}
        sheet_xml = contents[sheet_path].decode("utf-8")
        sheet_xml = sheet_xml.replace(
            '<row r="1"></row>',
            '<row r="1">'
            '<c r="A1"><f t="shared" ref="A1:A3" si="0">1+1</f><v>2</v></c>'
            "</row>",
        )
        contents[sheet_path] = sheet_xml.encode("utf-8")
        with zipfile.ZipFile(template, "w", zipfile.ZIP_DEFLATED) as zf:
            for name, data in contents.items():
                zf.writestr(name, data)

        skipped = apply_surgical_writes(
            template, output,
            [("Sheet1", "A", 1, 999), ("Sheet1", "C", 1, "hello")],
        )

        assert skipped == {("Sheet1", "A", 1)}

        with zipfile.ZipFile(output, "r") as zf:
            written_xml = zf.read(sheet_path).decode("utf-8")
        assert '<f t="shared" ref="A1:A3" si="0">1+1</f>' in written_xml
        assert "999" not in written_xml
        assert "hello" in written_xml
