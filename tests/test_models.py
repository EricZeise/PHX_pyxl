"""Tests for phpp_tool.models — Pydantic building record models."""

from __future__ import annotations

import json

import pytest

from phpp_tool.models import (
    BuildingRecord,
    ClimateData,
    ClimateDefinedRanges,
    DHWData,
    DHWPiping,
    DHWTankData,
    DHWTanks,
    OverviewData,
    OverviewBasicData,
    VerificationData,
    Windows,
    WindowRow,
)


# ---------------------------------------------------------------------------
# Verification
# ---------------------------------------------------------------------------


class TestVerification:
    def test_from_reader_data(self):
        data = {
            "setpoint_winter": 20,
            "setpoint_summer": 25,
            "num_of_units": 1,
            "phi_certification_type": "10-Passive House",
        }
        v = VerificationData(**data)
        assert v.setpoint_winter == 20
        assert v.setpoint_summer == 25
        assert v.num_of_units == 1
        assert v.phi_certification_type == "10-Passive House"

    def test_none_fields(self):
        v = VerificationData()
        assert v.setpoint_winter is None
        assert v.num_of_units is None

    def test_extra_fields_allowed(self):
        v = VerificationData(setpoint_winter=20, some_future_field="hello")
        assert v.setpoint_winter == 20
        assert v.some_future_field == "hello"


# ---------------------------------------------------------------------------
# Overview
# ---------------------------------------------------------------------------


class TestOverview:
    def test_basic_data(self):
        bd = OverviewBasicData(
            address_project_name="End-of-terrace Passive House ",
            address_number_dwellings_res=1,
            address_number_occupants_res=2.95,
        )
        assert bd.address_project_name == "End-of-terrace Passive House "
        assert bd.address_number_dwellings_res == 1

    def test_overview_with_sections(self):
        o = OverviewData(
            basic_data={"address_project_name": "Test House"},
            building_envelope={"address_area_envelope": 450.0},
        )
        assert o.basic_data.address_project_name == "Test House"
        assert o.building_envelope["address_area_envelope"] == 450.0


# ---------------------------------------------------------------------------
# Climate
# ---------------------------------------------------------------------------


class TestClimate:
    def test_defined_ranges(self):
        dr = ClimateDefinedRanges(
            climate_zone="3: Cool-temperate",
            latitude=51.301,
            longitude=9.44,
            site_altitude=None,
        )
        assert dr.climate_zone == "3: Cool-temperate"
        assert dr.latitude == 51.301
        assert dr.site_altitude is None

    def test_climate_model(self):
        c = ClimateData(
            defined_ranges={
                "climate_zone": "3: Cool-temperate", "latitude": 51.301},
            named_ranges={"country": "Europe"},
        )
        assert c.defined_ranges.climate_zone == "3: Cool-temperate"
        assert c.named_ranges["country"] == "Europe"


# ---------------------------------------------------------------------------
# DHW
# ---------------------------------------------------------------------------


class TestDHW:
    def test_piping(self):
        p = DHWPiping(
            total_length=13.5,
            diameter=20,
            insul_thickness=40,
            water_temp=60,
        )
        assert p.total_length == 13.5
        assert p.diameter == 20

    def test_tank_data(self):
        td = DHWTankData(
            tank_type="1-DHW and heating",
            standby_losses=1.2,
            storage_capacity=200,
        )
        assert td.tank_type == "1-DHW and heating"

    def test_tanks(self):
        t = DHWTanks(
            tank_1={"tank_type": "1-DHW and heating", "standby_losses": 1.2},
            tank_2={"tank_type": "0-No storage tank"},
        )
        assert t.tank_1.tank_type == "1-DHW and heating"
        assert t.tank_2.tank_type == "0-No storage tank"

    def test_dhw_from_reader(self):
        data = {
            "recirc_piping": {
                "total_length": 13.5,
                "diameter": 20},
            "tanks": {
                "tank_1": {
                    "tank_type": "1-DHW and heating",
                    "standby_losses": 1.2},
                "tank_2": {
                    "tank_type": "0-No storage tank"},
            },
        }
        dhw = DHWData(**data)
        assert dhw.recirc_piping.total_length == 13.5
        assert dhw.tanks.tank_1.tank_type == "1-DHW and heating"


# ---------------------------------------------------------------------------
# Windows
# ---------------------------------------------------------------------------


class TestWindows:
    def test_window_row(self):
        wr = WindowRow(quantity=2, description="Living room",
                       width=1.5, height=2.0)
        assert wr.quantity == 2
        assert wr.description == "Living room"

    def test_windows_with_dicts(self):
        w = Windows(window_rows=[
            {"quantity": 2, "description": "South", "_row": 11},
            {"quantity": 1, "description": "East", "_row": 12},
        ])
        assert len(w.window_rows) == 2


# ---------------------------------------------------------------------------
# BuildingRecord — full round-trip
# ---------------------------------------------------------------------------


class TestBuildingRecord:
    def test_from_reader_dict(self):
        data = {
            "VERIFICATION": {
                "setpoint_winter": 20,
                "setpoint_summer": 25,
                "num_of_units": 1,
            },
            "OVERVIEW": {
                "basic_data": {
                    "address_project_name": "Test House",
                    "address_number_dwellings_res": 1,
                },
            },
            "CLIMATE": {
                "defined_ranges": {
                    "climate_zone": "3: Cool-temperate",
                    "latitude": 51.301,
                },
            },
            "DHW": {
                "recirc_piping": {"total_length": 13.5},
                "tanks": {
                    "tank_1": {"tank_type": "1-DHW and heating"},
                },
            },
        }
        rec = BuildingRecord.from_reader_dict(data)
        assert rec.VERIFICATION.setpoint_winter == 20
        assert rec.OVERVIEW.basic_data.address_project_name == "Test House"
        assert rec.CLIMATE.defined_ranges.climate_zone == "3: Cool-temperate"
        assert rec.DHW.tanks.tank_1.tank_type == "1-DHW and heating"

    def test_to_json_round_trip(self):
        data = {
            "VERIFICATION": {"setpoint_winter": 20, "num_of_units": 1},
            "OVERVIEW": {
                "basic_data": {"address_project_name": "Test"},
            },
        }
        rec = BuildingRecord.from_reader_dict(data)
        json_str = rec.to_json()
        parsed = json.loads(json_str)

        assert parsed["VERIFICATION"]["setpoint_winter"] == 20
        basic = parsed["OVERVIEW"]["basic_data"]
        assert basic["address_project_name"] == "Test"

    def test_json_excludes_none(self):
        rec = BuildingRecord(
            VERIFICATION=VerificationData(setpoint_winter=20),
        )
        json_str = rec.to_json()
        parsed = json.loads(json_str)

        assert "VERIFICATION" in parsed
        assert "OVERVIEW" not in parsed
        assert "setpoint_summer" not in parsed["VERIFICATION"]

    def test_empty_record(self):
        rec = BuildingRecord()
        json_str = rec.to_json()
        parsed = json.loads(json_str)
        assert parsed == {}

    def test_unknown_worksheet_key(self):
        data = {
            "VERIFICATION": {"setpoint_winter": 20},
            "SOME_FUTURE_SHEET": {"field": "value"},
        }
        rec = BuildingRecord.from_reader_dict(data)
        assert rec.VERIFICATION.setpoint_winter == 20
        assert rec.SOME_FUTURE_SHEET == {"field": "value"}

    def test_config_key_in_worksheet(self):
        data = {
            "HEATING_DEMAND": {
                "_config": {
                    "unit": "kWh",
                    "col_kWh_year": "O",
                    "row_annual_demand": 78,
                },
            },
        }
        rec = BuildingRecord.from_reader_dict(data)
        assert rec.HEATING_DEMAND["_config"]["unit"] == "kWh"


class TestBuildingRecordIntegration:
    """Test that reader-shaped data validates through the model.

    Uses hand-crafted dicts (no Excel needed). The full reader→model
    roundtrip is covered by TestExcelRoundTrip in test_cli.py.
    """

    def test_reader_output_validates(self):
        data = {
            "VERIFICATION": {
                "setpoint_winter": 20,
                "num_of_units": 1,
            }
        }
        rec = BuildingRecord.from_reader_dict(data)

        assert rec.VERIFICATION.setpoint_winter == 20
        assert rec.VERIFICATION.num_of_units == 1

        json_str = rec.to_json()
        parsed = json.loads(json_str)
        assert parsed["VERIFICATION"]["setpoint_winter"] == 20

        rec2 = BuildingRecord.model_validate_json(json_str)
        assert rec2.VERIFICATION.setpoint_winter == 20
