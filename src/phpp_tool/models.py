"""Pydantic models for the PHPP building record.

These models validate the nested dict produced by ``reader.read_phpp()``
and provide a typed schema for JSON serialization/deserialization.

Design: core worksheets (Verification, Overview, Climate, Ventilation, DHW,
Windows) have explicit fields. Less-used or highly variable worksheets use
``dict[str, Any]`` to stay resilient to field map changes.
"""

from __future__ import annotations

from typing import Any

from pydantic import BaseModel, ConfigDict


class _Base(BaseModel):
    model_config = ConfigDict(extra="allow")


# ---------------------------------------------------------------------------
# Verification
# ---------------------------------------------------------------------------

class VerificationData(_Base):
    phi_building_category_type: Any = None
    phi_building_use_type: Any = None
    phi_building_ihg_type: Any = None
    phi_building_occupancy_type: Any = None
    phi_certification_type: Any = None
    phi_certification_class: Any = None
    phi_pe_type: Any = None
    phi_enerphit_type: Any = None
    phi_retrofit_type: Any = None
    num_of_units: int | float | None = None
    setpoint_winter: int | float | None = None
    setpoint_summer: int | float | None = None
    mechanical_cooling: Any = None


# ---------------------------------------------------------------------------
# Overview
# ---------------------------------------------------------------------------

class OverviewBasicData(_Base):
    address_project_name: str | None = None
    address_number_dwellings_res: int | float | None = None
    address_number_dwellings_nonres: int | float | None = None
    address_number_occupants_res: float | None = None
    address_number_occupants_nonres: Any = None


class OverviewData(_Base):
    basic_data: OverviewBasicData | None = None
    building_envelope: dict[str, Any] | None = None
    ventilation: dict[str, Any] | None = None


# ---------------------------------------------------------------------------
# Climate
# ---------------------------------------------------------------------------

class ClimateDefinedRanges(_Base):
    climate_zone: str | None = None
    weather_station_altitude: int | float | None = None
    site_altitude: int | float | None = None
    latitude: float | None = None
    longitude: float | None = None


class ClimateData(_Base):
    active_dataset: Any = None
    active_block: dict[str, Any] | None = None
    ud_block: Any = None
    named_ranges: dict[str, Any] | None = None
    defined_ranges: ClimateDefinedRanges | None = None


# ---------------------------------------------------------------------------
# Ventilation
# ---------------------------------------------------------------------------

class VentilationData(_Base):
    vent_type: Any = None
    wind_coeff_e: float | None = None
    wind_coeff_f: float | None = None
    airtightness_n50: float | None = None
    airtightness_Vn50: float | None = None
    multi_unit_on: Any = None


# ---------------------------------------------------------------------------
# Areas
# ---------------------------------------------------------------------------

class Areas(_Base):
    tfa_input: float | None = None
    summary_rows: dict[str, Any] | None = None
    surface_rows: list[dict[str, Any]] | None = None
    thermal_bridge_rows: list[dict[str, Any]] | None = None
    defined_ranges: dict[str, Any] | None = None


# ---------------------------------------------------------------------------
# Windows / Shading
# ---------------------------------------------------------------------------

class WindowRow(_Base):
    quantity: int | float | None = None
    description: str | None = None
    width: float | None = None
    height: float | None = None
    host: Any = None
    glazing_id: Any = None
    frame_id: Any = None
    orientation_angle: Any = None
    vertical_angle: Any = None
    orientation_label: Any = None


class Windows(_Base):
    window_rows: list[WindowRow] | list[dict[str, Any]] | None = None


class Shading(_Base):
    shading_rows: list[dict[str, Any]] | None = None


# ---------------------------------------------------------------------------
# Components
# ---------------------------------------------------------------------------

class Components(_Base):
    glazings: list[dict[str, Any]] | dict[str, Any] | None = None
    frames: list[dict[str, Any]] | None = None
    ventilators: list[dict[str, Any]] | None = None


# ---------------------------------------------------------------------------
# DHW
# ---------------------------------------------------------------------------

class DHWPiping(_Base):
    total_length: float | None = None
    diameter: float | None = None
    insul_thickness: float | None = None
    insul_reflective: Any = None
    insul_conductivity: float | None = None
    daily_period: float | None = None
    water_temp: float | None = None


class DHWTankData(_Base):
    tank_type: Any = None
    standby_losses: float | None = None
    storage_capacity: float | None = None
    standby_fraction: float | None = None
    tank_location: Any = None
    water_temp: float | None = None


class DHWTanks(_Base):
    tank_1: DHWTankData | None = None
    tank_2: DHWTankData | None = None
    tank_buffer: DHWTankData | None = None


class DHWData(_Base):
    recirc_piping: DHWPiping | dict[str, Any] | None = None
    branch_piping: DHWPiping | dict[str, Any] | None = None
    tanks: DHWTanks | dict[str, Any] | None = None


# ---------------------------------------------------------------------------
# Additional ventilation
# ---------------------------------------------------------------------------

class AddnlVent(_Base):
    rooms: list[dict[str, Any]] | None = None
    units: list[dict[str, Any]] | None = None
    ducts: list[dict[str, Any]] | None = None


# ---------------------------------------------------------------------------
# Top-level building record
# ---------------------------------------------------------------------------

class BuildingRecord(_Base):
    """Complete building record extracted from a PHPP workbook.

    Worksheet keys map to typed sub-models where available, or to
    ``dict[str, Any]`` for less-structured worksheets.
    """
    VERIFICATION: VerificationData | None = None
    OVERVIEW: OverviewData | None = None
    CLIMATE: ClimateData | None = None
    VENTILATION: VentilationData | None = None
    AREAS: Areas | None = None
    UVALUES: dict[str, Any] | None = None
    COMPONENTS: Components | None = None
    WINDOWS: Windows | None = None
    SHADING: Shading | None = None
    ADDNL_VENT: AddnlVent | None = None
    DHW: DHWData | None = None
    HEATING_DEMAND: dict[str, Any] | None = None
    HEATING_PEAK_LOAD: dict[str, Any] | None = None
    COOLING_DEMAND: dict[str, Any] | None = None
    COOLING_PEAK_LOAD: dict[str, Any] | None = None
    COOLING_UNITS: dict[str, Any] | None = None
    SOLAR_DHW: dict[str, Any] | None = None
    SOLAR_PV: dict[str, Any] | None = None
    ELECTRICITY: dict[str, Any] | None = None
    ELEC_NON_RES: dict[str, Any] | None = None
    PER: dict[str, Any] | None = None
    VARIANTS: dict[str, Any] | None = None
    DATA: dict[str, Any] | None = None
    SUMM_VENT: dict[str, Any] | None = None
    USE_NON_RES: dict[str, Any] | None = None
    AUX_ELEC: dict[str, Any] | None = None
    IHG_NON_RES: dict[str, Any] | None = None
    HP: dict[str, Any] | None = None
    BOILER: dict[str, Any] | None = None
    GROUND: dict[str, Any] | None = None
    EASY_PH: dict[str, Any] | None = None

    def to_json(self, **kwargs: Any) -> str:
        return self.model_dump_json(indent=2, exclude_none=True, **kwargs)

    @classmethod
    def from_reader_dict(cls, data: dict[str, Any]) -> BuildingRecord:
        """Create from the dict returned by read_phpp()."""
        return cls.model_validate(data)
