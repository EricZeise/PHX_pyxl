"""Tests for the field map parser against phpp-field-mapping.md."""

from pathlib import Path

import pytest

from phpp_tool.map_parser import parse_field_map

FIELD_MAP = Path(__file__).resolve().parent.parent / "phpp-field-mapping.md"


@pytest.fixture(scope="module")
def field_map():
    return parse_field_map(FIELD_MAP)


# ---------------------------------------------------------------------------
# Structural completeness
# ---------------------------------------------------------------------------

EXPECTED_KEYS = [
    "VERIFICATION", "VARIANTS", "CLIMATE", "UVALUES", "AREAS",
    "GROUND", "COMPONENTS", "WINDOWS", "SHADING", "VENTILATION",
    "ADDNL_VENT", "HEATING_DEMAND", "HEATING_PEAK_LOAD",
    "COOLING_DEMAND", "COOLING_PEAK_LOAD", "SUMM_VENT",
    "COOLING_UNITS", "DHW", "SOLAR_DHW", "SOLAR_PV",
    "ELECTRICITY", "ELEC_NON_RES", "USE_NON_RES", "AUX_ELEC",
    "IHG_NON_RES", "PER", "HP", "BOILER", "DATA", "OVERVIEW",
    "EASY_PH",
]


def test_all_worksheet_keys_present(field_map):
    for key in EXPECTED_KEYS:
        assert key in field_map, f"Missing worksheet key: {key}"


def test_sheet_names(field_map):
    assert field_map["VERIFICATION"]["sheet_name"] == "Verification"
    assert field_map["CLIMATE"]["sheet_name"] == "Climate"
    assert field_map["WINDOWS"]["sheet_name"] == "Windows"
    assert field_map["DHW"]["sheet_name"] == "DHW+Distribution"
    assert field_map["OVERVIEW"]["sheet_name"] == "Overview"
    assert field_map["EASY_PH"]["sheet_name"] == "easyPH"


def test_stubs_parse_without_error(field_map):
    for key in ["GROUND", "SUMM_VENT", "USE_NON_RES", "AUX_ELEC",
                "IHG_NON_RES", "HP", "BOILER", "EASY_PH"]:
        ws = field_map[key]
        assert ws["worksheet_key"] == key


# ---------------------------------------------------------------------------
# Strategy 1: Label-anchored relative
# ---------------------------------------------------------------------------

class TestLabelAnchored:
    def test_setpoint_winter(self, field_map):
        f = field_map["VERIFICATION"]["fields"]["setpoint_winter"]
        assert f["locator_string"] == "Interior temperature winter [°C]:"
        assert f["locator_col"] == "J"
        assert f["input_col"] == "K"
        assert f["row_offset"] == 0
        assert f["unit"] == "C"
        assert f["options"] is None

    def test_setpoint_summer(self, field_map):
        f = field_map["VERIFICATION"]["fields"]["setpoint_summer"]
        assert f["locator_string"] == "Interior temp. summer [°C]:"
        assert f["locator_col"] == "M"
        assert f["input_col"] == "N"
        assert f["row_offset"] == 0
        assert f["unit"] == "C"

    def test_num_of_units(self, field_map):
        f = field_map["VERIFICATION"]["fields"]["num_of_units"]
        assert f["locator_string"] == "No. of dwelling units:"
        assert f["locator_col"] == "E"
        assert f["input_col"] == "F"
        assert f["row_offset"] == 0

    def test_phi_certification_type_with_options(self, field_map):
        f = field_map["VERIFICATION"]["fields"]["phi_certification_type"]
        assert f["locator_string"] == "Planned energy standard"
        assert f["locator_col"] == "T"
        assert f["input_col"] == "T"
        assert f["row_offset"] == 1
        assert f["options"] is not None
        assert "10" in f["options"]
        assert "10-Passive house" in f["options"]["10"]

    def test_phi_building_category_options(self, field_map):
        f = field_map["VERIFICATION"]["fields"]["phi_building_category_type"]
        assert f["options"] is not None
        assert len(f["options"]) == 4

    def test_phi_certification_class_malformed_row(self, field_map):
        """The phi_certification_class row has an extra column split."""
        f = field_map["VERIFICATION"]["fields"]["phi_certification_class"]
        assert f["locator_col"] == "T"
        assert f["input_col"] == "T"
        assert f["row_offset"] == 1
        assert f["options"] is not None
        assert "10" in f["options"]
        assert "20" in f["options"]

    def test_ventilation_label_fields(self, field_map):
        f = field_map["VENTILATION"]["fields"]["vent_type"]
        assert f["locator_string"] == "Type of ventilation"
        assert f["locator_col"] == "I"
        assert f["input_col"] == "K"

    def test_areas_tfa_input(self, field_map):
        f = field_map["AREAS"]["fields"]["tfa_input"]
        assert f["locator_string"] == "1-Treated floor area"
        assert f["locator_col"] == "M"
        assert f["input_col"] == "T"
        assert f["unit"] == "M2"


# ---------------------------------------------------------------------------
# Strategy 2: Header + entry locator (repeating blocks)
# ---------------------------------------------------------------------------

class TestHeaderEntryBlock:
    def test_window_rows(self, field_map):
        sec = field_map["WINDOWS"]["sections"]["window_rows"]
        assert sec["header_locator"] == {
            "col": "L", "string": "Windows and entrance doors"
        }
        assert sec["entry_locator"] == {"col": "L", "string": "Quan-"}
        assert "quantity" in sec["column_fields"]
        assert sec["column_fields"]["quantity"]["column"] == "L"
        assert sec["column_fields"]["width"]["column"] == "R"
        assert sec["column_fields"]["width"]["unit"] == "M"

    def test_window_rows_end(self, field_map):
        sec = field_map["WINDOWS"]["sections"]["window_rows_end"]
        assert sec["entry_locator"]["string"] == "Unhide additional rows"

    def test_shading_rows(self, field_map):
        sec = field_map["SHADING"]["sections"]["shading_rows"]
        assert sec["header_locator"]["col"] == "T"
        assert "h_hori" in sec["column_fields"]
        assert sec["column_fields"]["h_hori"]["unit"] == "M"

    def test_areas_surface_rows(self, field_map):
        sec = field_map["AREAS"]["sections"]["surface_rows"]
        assert sec["header_locator"]["string"] == "Area input"
        assert sec["entry_locator"]["string"] == "1"
        assert sec["column_fields"]["area"]["column"] == "T"
        assert sec["column_fields"]["area"]["unit"] == "M2"

    def test_areas_thermal_bridge_rows(self, field_map):
        sec = field_map["AREAS"]["sections"]["thermal_bridge_rows"]
        assert "psi_value" in sec["column_fields"]

    def test_components_frames(self, field_map):
        sec = field_map["COMPONENTS"]["sections"]["frames"]
        assert sec["header_locator"]["string"] == "Window and door frames"
        assert sec["entry_locator"]["string"] == "01ud"
        assert sec["column_fields"]["psi_g_left"]["column"] == "IR"
        assert sec["column_fields"]["psi_g_right"]["column"] == "IR"

    def test_components_ventilators(self, field_map):
        sec = field_map["COMPONENTS"]["sections"]["ventilators"]
        assert sec["header_locator"]["col"] == "LQ"
        assert "sensible_heat_recovery" in sec["column_fields"]

    def test_addnl_vent_rooms(self, field_map):
        sec = field_map["ADDNL_VENT"]["sections"]["rooms"]
        assert sec["header_locator"]["string"] == "Room"
        assert "weighted_floor_area" in sec["column_fields"]

    def test_addnl_vent_ducts(self, field_map):
        sec = field_map["ADDNL_VENT"]["sections"]["ducts"]
        assert sec["column_fields"]["duct_assign_9"]["column"] == "Z"
        assert sec["column_fields"]["duct_assign_10"]["column"] == "Z"

    def test_elec_nonres_lighting_rows(self, field_map):
        sec = field_map["ELEC_NON_RES"]["sections"]["lighting_rows"]
        assert sec["header_locator"]["string"] == "Lighting"
        assert sec["entry_locator"]["string"] == "Room / Zone"
        assert "installed_power" in sec["column_fields"]
        assert sec["column_fields"]["installed_power"]["unit"] == "W/M2"


# ---------------------------------------------------------------------------
# Strategy 3: Named ranges
# ---------------------------------------------------------------------------

class TestNamedRanges:
    def test_climate_named_ranges(self, field_map):
        sec = field_map["CLIMATE"]["sections"]["named_ranges"]
        assert sec["items"]["country"] == "Klima_Region"
        assert sec["items"]["region"] == "Klima_Region2"
        assert sec["items"]["data_set"] == "Klima_Standort"

    def test_cooling_units_supply_air_named_ranges(self, field_map):
        sec = field_map["COOLING_UNITS"]["sections"]["supply_air"]
        assert sec["items"]["used"] == "Kuehlgeraete_Zuluft_Kuehlung_Ankreuzen"
        assert (sec["items"]["num_units"]
                == "Kuehlgeraete_Kompressor_Zuluft_Anzahl")

    def test_per_named_ranges(self, field_map):
        sec = field_map["PER"]["sections"]["named_ranges"]
        assert sec["items"]["heating_type_1"] == "PE_Waermeerzeuger_primaer"

    def test_solar_dhw_footprint_named_range(self, field_map):
        sec = field_map["SOLAR_DHW"]["sections"]["ranges"]
        assert sec["items"]["footprint"] == "SolarWW_Kollektorflaeche"


# ---------------------------------------------------------------------------
# Strategy 4: Absolute addresses
# ---------------------------------------------------------------------------

class TestAbsoluteAddresses:
    def test_climate_defined_ranges(self, field_map):
        sec = field_map["CLIMATE"]["sections"]["defined_ranges"]
        assert sec["items"]["climate_zone"] == "D13"
        assert sec["items"]["latitude"] == "F25"
        assert sec["items"]["longitude"] == "H25"
        assert sec["items"]["site_altitude"] == "D18"

    def test_overview_basic_data_addresses(self, field_map):
        sec = field_map["OVERVIEW"]["sections"]["basic_data"]
        assert sec["items"]["address_project_name"] == "C11"
        assert sec["items"]["address_number_dwellings_res"] == "C27"
        assert sec["items"]["address_number_occupants_res"] == "C28"

    def test_overview_building_envelope_col_row(self, field_map):
        sec = field_map["OVERVIEW"]["sections"]["building_envelope"]
        addr = sec["items"]["address_area_envelope"]
        assert addr == {"col": "C", "row": 105, "unit": "M2"}

    def test_overview_ventilation_col_row(self, field_map):
        sec = field_map["OVERVIEW"]["sections"]["ventilation"]
        addr = sec["items"]["vn50"]
        assert addr == {"col": "C", "row": 362, "unit": "M3"}

    def test_cooling_demand_absolute_addresses(self, field_map):
        cfg = field_map["COOLING_DEMAND"]["config"]
        assert cfg["address_specific_latent_cooling_demand"] == "AN176"
        assert cfg["address_tfa"] == "O8"

    def test_solar_dhw_absolute_addresses(self, field_map):
        sec = field_map["SOLAR_DHW"]["sections"]["ranges"]
        assert sec["items"]["annual_dhw_contribution"] == "N34"


# ---------------------------------------------------------------------------
# Strategy 5: Column + row-offset within a block
# ---------------------------------------------------------------------------

class TestRowOffsetBlock:
    def test_uvalues_constructor_header(self, field_map):
        sec = field_map["UVALUES"]["sections"]["constructor"]
        assert sec["header_locator"] == {
            "col": "L",
            "string": "Description of building assembly",
        }

    def test_uvalues_constructor_columns(self, field_map):
        sec = field_map["UVALUES"]["sections"]["constructor"]
        assert sec["column_fields"]["display_name"]["column"] == "L"
        assert sec["column_fields"]["r_si"]["column"] == "M"
        assert sec["column_fields"]["thickness"]["column"] == "R"
        assert sec["column_fields"]["thickness"]["unit"] == "MM"

    def test_uvalues_constructor_row_offsets(self, field_map):
        sec = field_map["UVALUES"]["sections"]["constructor"]
        items = sec["items"]
        assert items["name_row_offset"] == 2
        assert items["rsi_row_offset"] == 4
        assert items["rse_row_offset"] == 5
        assert items["first_layer_row_offset"] == 7
        assert items["last_layer_row_offset"] == 14
        assert items["result_val_row_offset"] == 19
        assert items["result_val_col"] == "R"
        assert items["result_val_unit"] == "W/M2K"

    def test_dhw_recirc_piping(self, field_map):
        sec = field_map["DHW"]["sections"]["recirc_piping"]
        assert sec["header_locator"]["string"] == "DHW distribution"
        assert sec["entry_locator"]["string"] == (
            "DHW circulation pipes or, for heat interface units, "
            "forward and return flows"
        )
        assert sec["row_fields"]["total_length"]["row_offset"] == 2
        assert sec["row_fields"]["total_length"]["unit"] == "M"
        assert sec["row_fields"]["water_temp"]["row_offset"] == 13
        assert sec["row_fields"]["water_temp"]["unit"] == "C"

    def test_dhw_branch_piping(self, field_map):
        sec = field_map["DHW"]["sections"]["branch_piping"]
        assert sec["row_fields"]["diameter"]["row_offset"] == 2
        assert sec["row_fields"]["diameter"]["unit"] == "MM"


# ---------------------------------------------------------------------------
# Strategy 6: Fixed result rows/cols
# ---------------------------------------------------------------------------

class TestFixedRowsCols:
    def test_heating_demand_config(self, field_map):
        cfg = field_map["HEATING_DEMAND"]["config"]
        assert cfg["col_kWh_year"] == "O"
        assert cfg["col_kWh_m2_year"] == "Q"
        assert cfg["row_total_losses_transmission"] == 27
        assert cfg["row_annual_demand"] == 78
        assert cfg["unit"] == "kWh"

    def test_heating_peak_load_config(self, field_map):
        cfg = field_map["HEATING_PEAK_LOAD"]["config"]
        assert cfg["col_weather_1"] == "P"
        assert cfg["row_total_load"] == 88

    def test_cooling_demand_config(self, field_map):
        cfg = field_map["COOLING_DEMAND"]["config"]
        assert cfg["row_annual_sensible_demand"] == 88
        assert cfg["row_annual_latent_demand"] == 94

    def test_cooling_peak_load_config(self, field_map):
        cfg = field_map["COOLING_PEAK_LOAD"]["config"]
        assert cfg["row_total_sensible_load"] == 64
        assert cfg["row_total_latent_load"] == 93


# ---------------------------------------------------------------------------
# DHW tanks: mixed column + row + entry_row_start
# ---------------------------------------------------------------------------

class TestDHWTanks:
    def test_tanks_locators(self, field_map):
        sec = field_map["DHW"]["sections"]["tanks"]
        assert sec["header_locator"]["string"] == "Storage heat losses"
        assert sec["entry_locator"]["string"] == "Storage type 1"

    def test_tanks_entry_row_start(self, field_map):
        sec = field_map["DHW"]["sections"]["tanks"]
        assert sec["items"]["entry_row_start"] == 191

    def test_tanks_column_fields(self, field_map):
        sec = field_map["DHW"]["sections"]["tanks"]
        assert sec["column_fields"]["tank_1"]["column"] == "J"
        assert sec["column_fields"]["tank_2"]["column"] == "M"
        assert sec["column_fields"]["tank_buffer"]["column"] == "P"

    def test_tanks_row_fields(self, field_map):
        sec = field_map["DHW"]["sections"]["tanks"]
        assert sec["row_fields"]["tank_type"]["row"] == 0
        assert sec["row_fields"]["standby_losses"]["row"] == 5
        assert sec["row_fields"]["water_temp"]["row"] == 12


# ---------------------------------------------------------------------------
# Components glazings: fixed start row
# ---------------------------------------------------------------------------

class TestComponentsGlazings:
    def test_glazings_items(self, field_map):
        sec = field_map["COMPONENTS"]["sections"]["glazings"]
        assert sec["items"]["entry_column"] == "IH"
        assert sec["items"]["entry_start_row"] == 13
        assert sec["items"]["header_start_row"] == 8

    def test_glazings_columns(self, field_map):
        sec = field_map["COMPONENTS"]["sections"]["glazings"]
        assert sec["column_fields"]["g_value"]["column"] == "IJ"
        assert sec["column_fields"]["u_value"]["unit"] == "W/M2K"


# ---------------------------------------------------------------------------
# Variants: header locators + input items
# ---------------------------------------------------------------------------

class TestVariants:
    def test_variants_config(self, field_map):
        cfg = field_map["VARIANTS"]["config"]
        assert cfg["active_value_column"] == "E"

    def test_variants_ventilation_section(self, field_map):
        sec = field_map["VARIANTS"]["sections"]["ventilation"]
        assert sec["header_locator"]["string"] == "Ventilation"
        assert sec["items"]["input_col"] == "C"

    def test_variants_assemblies_section(self, field_map):
        sec = field_map["VARIANTS"]["sections"]["assemblies"]
        assert sec["header_locator"]["string"] == "Building assembly layers"


# ---------------------------------------------------------------------------
# Climate: complex multi-part sections
# ---------------------------------------------------------------------------

class TestClimate:
    def test_active_block_items(self, field_map):
        sec = field_map["CLIMATE"]["sections"]["active_block"]
        assert sec["items"]["start_row"] == 26
        assert sec["items"]["end_row"] == 36
        assert sec["items"]["start_col"] == "D"
        assert sec["items"]["end_col"] == "U"

    def test_ud_block_column_fields(self, field_map):
        sec = field_map["CLIMATE"]["sections"]["ud_block"]
        assert sec["column_fields"]["jan"]["column"] == "E"
        assert sec["column_fields"]["dec"]["column"] == "P"
        assert sec["column_fields"]["latitude"]["column"] == "F"

    def test_ud_block_row_fields(self, field_map):
        sec = field_map["CLIMATE"]["sections"]["ud_block"]
        assert sec["row_fields"]["temperature_air"]["row"] == 1
        assert sec["row_fields"]["temperature_air"]["unit"] == "C"
        assert sec["row_fields"]["radiation_south"]["row"] == 4


# ---------------------------------------------------------------------------
# Electricity: appliance rows
# ---------------------------------------------------------------------------

class TestElectricity:
    def test_input_columns(self, field_map):
        sec = field_map["ELECTRICITY"]["sections"]["input_columns"]
        assert sec["items"]["selection"] == "E"
        assert sec["items"]["annual_energy_demand"] == "AB"

    def test_input_rows_appliances(self, field_map):
        sec = field_map["ELECTRICITY"]["sections"]["input_rows"]
        assert sec["appliance_rows"]["dishwasher"]["data_row"] == 25
        assert sec["appliance_rows"]["dishwasher"]["selection_row"] == 24
        assert "1" in sec["appliance_rows"]["dishwasher"]["options"]

    def test_cooking_options(self, field_map):
        sec = field_map["ELECTRICITY"]["sections"]["input_rows"]
        opts = sec["appliance_rows"]["cooking"]["options"]
        assert opts["1"] == "1-Electricity"
        assert opts["3"] == "3-LPG"

    def test_refrigerator_no_selection_row(self, field_map):
        sec = field_map["ELECTRICITY"]["sections"]["input_rows"]
        assert "selection_row" not in sec["appliance_rows"]["refrigerator"]


# ---------------------------------------------------------------------------
# PER: heading-based sections
# ---------------------------------------------------------------------------

class TestPER:
    def test_per_config(self, field_map):
        cfg = field_map["PER"]["config"]
        assert cfg["locator_col"] == "P"
        assert cfg["unit"] == "KWH"

    def test_per_columns(self, field_map):
        sec = field_map["PER"]["sections"]["columns"]
        assert sec["items"]["calculated_efficiency"] == "Q"
        assert sec["items"]["co2_emissions"] == "Z"

    def test_per_heating(self, field_map):
        sec = field_map["PER"]["sections"]["heating"]
        assert sec["items"]["locator_string_heading"] == "Heating"
        assert (sec["items"]["locator_string_start"]
                == "Electricity (HP compact unit)")

    def test_per_addresses(self, field_map):
        sec = field_map["PER"]["sections"]["addresses"]
        assert sec["items"]["tfa"] == "Z7"
        assert sec["items"]["footprint"] == "Z8"


# ---------------------------------------------------------------------------
# Solar PV: columns/rows grid
# ---------------------------------------------------------------------------

class TestSolarPV:
    def test_pv_columns(self, field_map):
        sec = field_map["SOLAR_PV"]["sections"]["columns"]
        assert sec["items"]["systems_start"] == "S"
        assert sec["items"]["systems_end"] == "W"

    def test_pv_rows(self, field_map):
        sec = field_map["SOLAR_PV"]["sections"]["rows"]
        assert sec["items"]["systems_start"] == 10
        assert sec["items"]["annual_energy"] == 42


# ---------------------------------------------------------------------------
# Data: version section
# ---------------------------------------------------------------------------

def test_data_version(field_map):
    sec = field_map["DATA"]["sections"]["version"]
    assert sec["header_locator"]["string"] == "PHPP Version"
