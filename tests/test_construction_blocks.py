"""Tests for the opaque-construction ("R-Values"/"U-Values" grid) reader
and writer -- the repeating fixed-pitch row-block shape confirmed against
a real workbook 2026-09-22 (see EN_10_6_IP.md's own "Repeating-block
geometry" note), distinct from every other worksheet shape in this
project and so exercised here directly rather than only through the
full-file roundtrip scripts."""

from __future__ import annotations

from openpyxl import Workbook

from phpp_tool.locators import predicted_phpp_id_num
from phpp_tool.reader import _read_construction_blocks
from phpp_tool.writer import _write_construction_blocks

_ITEMS = {
    "block_pitch_rows": 21,
    "phpp_id_num_col_offset": 5,
    "name_row_offset": 2,
    "rsi_row_offset": 4,
    "rse_row_offset": 5,
    "first_layer_row_offset": 7,
    "last_layer_row_offset": 14,
    "percentage_row_offset": 15,
    "result_val_row_offset": 19,
    "result_val_col": "R",
}
_COLUMN_FIELDS = {
    "display_name": {"column": "L"},
    "r_si": {"column": "M"},
    "r_se": {"column": "M"},
    "interior_insulation": {"column": "R"},
    "u_val_supplement": {"column": "R"},
    "sec_1_description": {"column": "L"},
    "sec_1_conductivity": {"column": "M"},
    "sec_2_description": {"column": "N"},
    "sec_2_conductivity": {"column": "O"},
    "sec_3_description": {"column": "P"},
    "sec_3_conductivity": {"column": "Q"},
    "thickness": {"column": "R"},
    "sec_1_percentage": {"column": "M"},
    "sec_2_percentage": {"column": "O"},
    "sec_3_percentage": {"column": "Q"},
}
_HEADER_LOCATOR = {"col": "L", "string": "Description of building assembly"}
_SEC_SPEC = {
    "header_locator": _HEADER_LOCATOR,
    "items": _ITEMS,
    "column_fields": _COLUMN_FIELDS,
}


def test_predicted_phpp_id_num_matches_the_real_phpp_formula():
    # =IF((ROW()-ROW($Q$8))/21+1<10,0,"")&(ROW()-ROW($Q$8))/21+1&"ud"
    assert predicted_phpp_id_num(0) == "01ud"
    assert predicted_phpp_id_num(8) == "09ud"
    assert predicted_phpp_id_num(9) == "10ud"
    assert predicted_phpp_id_num(24) == "25ud"


def _build_two_block_sheet() -> Workbook:
    """Two real-shaped construction blocks, 21 rows apart, matching a real
    workbook's own layout (row/column values taken from a real IP-shell
    "R-Values" sheet, 2026-09-22)."""
    wb = Workbook()
    ws = wb.active
    # Block 1: "Externl wall", header at row 7.
    ws["L7"] = "Description of building assembly"
    ws["L8"], ws["Q8"] = "Externl wall", "01ud"
    ws["M10"] = "2-Wall"
    ws["M11"] = "1-Outdoor air"
    ws["L13"], ws["M13"], ws["R13"] = "Gypsum", 0.9, 0.625
    ws["L14"], ws["M14"] = "Cellulose", 3.6
    ws["N14"], ws["O14"], ws["R14"] = '2x4 @ 24"', 1.3, 3.5
    ws["O21"] = 0.094  # Percentage of sec. 2 (sec. 1's own M21 is a live
    # PHPP formula in the real sheet -- never modeled as writable here).

    # Block 2: "Roof", header at row 28 (7 + 21).
    ws["L28"] = "Description of building assembly"
    ws["L29"], ws["Q29"] = "Roof", "02ud"
    ws["M31"] = "1-Roof"
    ws["M32"] = "1-Outdoor air"
    ws["L34"], ws["M34"], ws["R34"] = "Gypsum", 0.9, 0.625
    return wb


class TestReadConstructionBlocks:
    def test_reads_two_real_shaped_blocks(self):
        wb = _build_two_block_sheet()
        ws = wb.active
        blocks = _read_construction_blocks((ws, ws), _SEC_SPEC, skip_formulas=True)

        assert len(blocks) == 2
        assert blocks[0]["display_name"] == "Externl wall"
        assert blocks[0]["phpp_id_num"] == "01ud"
        assert blocks[0]["r_si"] == "2-Wall"
        assert len(blocks[0]["layers"]) == 2
        assert blocks[0]["layers"][0]["sec_1_description"] == "Gypsum"
        assert blocks[0]["layers"][0]["thickness"] == 0.625
        assert blocks[0]["layers"][1]["sec_2_description"] == '2x4 @ 24"'
        assert blocks[0]["sec_2_percentage"] == 0.094

        assert blocks[1]["display_name"] == "Roof"
        assert blocks[1]["phpp_id_num"] == "02ud"
        assert len(blocks[1]["layers"]) == 1

    def test_a_stray_thickness_value_alone_is_not_a_real_layer(self):
        # A blank layer row that happens to carry a leftover literal 0 in
        # its own thickness cell (real, seen in a real workbook) must not
        # be reported as a real layer -- only a real material name does.
        wb = _build_two_block_sheet()
        ws = wb.active
        ws["R15"] = 0
        blocks = _read_construction_blocks((ws, ws), _SEC_SPEC, skip_formulas=True)
        assert len(blocks[0]["layers"]) == 2

    def test_stops_after_a_run_of_blank_block_headers(self):
        wb = _build_two_block_sheet()
        ws = wb.active
        # A third, real block far below a long run of blank ones (4 blank
        # block positions: rows 49, 70, 91, 112) -- never reached once the
        # blank-streak threshold (3) trips.
        ws["L133"] = "Description of building assembly"
        ws["L134"] = "Never reached"
        blocks = _read_construction_blocks((ws, ws), _SEC_SPEC, skip_formulas=True)
        assert [b["display_name"] for b in blocks] == ["Externl wall", "Roof"]


class TestWriteConstructionBlocks:
    def test_writes_scalar_layer_and_percentage_fields_at_the_right_cells(self):
        wb = Workbook()
        ws = wb.active
        ws["L7"] = "Description of building assembly"
        blocks_data = [
            {
                "display_name": "Externl wall",
                "r_si": "2-Wall",
                "r_se": "1-Outdoor air",
                "layers": [
                    {
                        "sec_1_description": "Gypsum",
                        "sec_1_conductivity": 0.9,
                        "thickness": 0.625,
                    },
                ],
                "sec_2_percentage": 0.094,
            },
            {
                "display_name": "Roof",
                "layers": [],
            },
        ]
        pending: list[tuple[str, str, int, object]] = []
        count = _write_construction_blocks(
            ws, "R-Values", _SEC_SPEC, blocks_data, pending
        )

        assert count == len(pending)
        assert ("R-Values", "L", 8, "Externl wall") in pending
        assert ("R-Values", "M", 10, "2-Wall") in pending
        assert ("R-Values", "M", 11, "1-Outdoor air") in pending
        assert ("R-Values", "L", 13, "Gypsum") in pending
        assert ("R-Values", "M", 13, 0.9) in pending
        assert ("R-Values", "R", 13, 0.625) in pending
        assert ("R-Values", "O", 21, 0.094) in pending
        # Second block, one full block_pitch_rows (21) below the first.
        assert ("R-Values", "L", 29, "Roof") in pending
        # Never a phpp_id_num or sec_1_percentage write -- both live
        # PHPP formulas.
        assert not any(w[1] == "Q" and w[2] in (8, 29) for w in pending)
        assert not any(w[1] == "M" and w[2] in (21, 42) for w in pending)

    def test_more_layers_than_phpp_reserves_are_truncated_with_a_warning(self, caplog):
        wb = Workbook()
        ws = wb.active
        ws["L7"] = "Description of building assembly"
        too_many_layers = [
            {
                "sec_1_description": f"Layer {i}",
                "sec_1_conductivity": 1.0,
                "thickness": 1.0,
            }
            for i in range(10)
        ]
        pending: list[tuple[str, str, int, object]] = []
        _write_construction_blocks(
            ws,
            "R-Values",
            _SEC_SPEC,
            [{"display_name": "Overstuffed", "layers": too_many_layers}],
            pending,
        )
        layer_rows = {w[2] for w in pending if w[1] == "L" and w[3].startswith("Layer")}
        assert len(layer_rows) == 8  # first_layer_row_offset..last_layer_row_offset
        assert "more than the 8" in caplog.text
