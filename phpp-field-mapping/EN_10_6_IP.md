# PHPP Field Mapping Reference — EN_10_6_IP

Localization map for **PHPP 10.6 (English, IP units)**. Targets an IP-shell PHPP workbook that also carries `<Name> SI`-suffixed mirror tabs (formula passthroughs of the base tab, used for SI-unit display only) — verified against `Data/Example_IP.xlsx`/`Data/Empty_IP.xlsx`. Sheet names are the same base (non-suffixed) names as EN_10_6_SI — reading/writing always targets the base tab directly, never the `<Name> SI` mirror, since the mirror tabs are formula-driven passthroughs and would silently lose designer-entered input under `skip_formulas`. The only content difference from EN_10_6_SI is locator strings for the handful of fields whose label text embeds a unit marker (`[°F]` vs `[°C]`, confirmed via direct verification against both file types); everything else is identical.

This document describes how PHX locates and reads/writes cells in each PHPP worksheet.

!!! note "Locator pattern"
    PHX locates fields by searching for `locator_string` in `locator_col`, then reads/writes at `input_column` offset by `input_row_offset` rows from the found row.

---

## Verification

Worksheet key: `VERIFICATION`

| Field Key | Label (`locator_string`) | Locator Col | Input Col | Row Offset | Unit | Options |
|-----------|-------------------------|-------------|-----------|------------|------|---------|
| `phi_building_category_type` | Building use | R | T | 1 |  | `1`: 10-Residential building: Residential; `2`: 12-Residential building: Other; `11`: 10-Residential building: Residential; `12`: 12-Residential building: Other |
| `phi_building_use_type` | Building use | R | T | 1 |  | `10`: 10-Residential building: Residential; `12`: 12-Residential building: Other; `20`: 20-Non-res building: Office/Administration; `21`: 21-Non-res building: School half-days (< 7 h); `22`: 22-Non-res building: School full-time (≥ 7 h); `23`: 23-Non-res.: Other |
| `phi_building_ihg_type` | Building use | R | T | 4 |  | `2`: 2-Standard; `3`: 3-PHPP-calculation ('IHG' worksheet); `4`: 4-PHPP-calculation ('IHG non-res' worksheet) |
| `phi_building_occupancy_type` | No. of occupants | R | R | 2 |  | `1`: *(null)*; `2`: *(null)* |
| `phi_certification_type` | Planned energy standard | T | T | 1 |  | `10`: 10-Passive house; `21`: 21-EnerPHit (Component method); `22`: 22-EnerPHit (Energy demand method); `30`: 30-PHI Low Energy Building; `40`: 40-Other |
| `phi_certification_class` | Class | Primary energy method | T | T | 1 |  | `10`: 10-Classic | PER (renewable); `11`: 11-Classic | PE (non-renewable); `20`: 20-Plus | PER (renewable); `30`: 30-Premium | PER (renewable) |
| `phi_pe_type` | Primary energy demand criterion | T | T | 1 |  | `1`: 1-Standard; `2`: 2-Project-specific |
| `phi_enerphit_type` | New building / Retrofit | T | T | 1 |  | `1`: 1-New building; `2`: 2-Retrofit; `3`: 3-Staged retrofit |
| `phi_retrofit_type` | New building / Retrofit | T | T | 1 |  | `1`: 1-New building; `2`: 2-Retrofit; `3`: 3-Staged retrofit |
| `num_of_units` | No. of dwelling units: | E | F | 0 |  |  |
| `setpoint_winter` | Interior temperature winter [°F]: | J | K | 0 | F |  |
| `setpoint_summer` | Interior temp. summer [°F]: | M | N | 0 | F |  |
| `mechanical_cooling` | Mechanical cooling: | M | N | 0 |  |  |

---

## Variants

Worksheet key: `VARIANTS`

**Configuration:**

- `active_value_column` (literal): `E`

### results_header

- Header locator: col `B`, string `"Results"`

### input_header

- Header locator: col `B`, string `"Input variables"`

### assemblies

- Header locator: col `C`, string `"Building assembly layers"`

- `input_col` (literal): `C`

### windows

- Header locator: col `C`, string `"Windows and shading"`

- `input_col` (literal): `C`

### ventilation

- Header locator: col `C`, string `"Ventilation"`

**Input item names:**

- `vent_type` (literal): "Ventilation type"
- `air_change_rate` (literal): "Air change rate at pressurisation test (n50)"
- `design_flow_rate` (literal): "Design air flow rate (maximum)"
- `install_location` (literal): "Installation site ventilation unit"
- `ventilator_unit` (literal): "Ventilation unit selection"

- `input_col` (literal): `C`

---

## Climate

Worksheet key: `CLIMATE`

### active_dataset

| Field | Column |
|-------|--------|
| `country` | D |
| `region` | D |
| `dataset` | D |
| `elevation_override` | D |

### active_block

- `start_row` (literal): `26`
- `end_row` (literal): `36`
- `start_col` (literal): `D`
- `end_col` (literal): `U`

### ud_block

- Header locator: col `Name of location`, string `""`
- `start_row` (literal): `67`

| Field | Column |
|-------|--------|
| `jan` | E |
| `feb` | F |
| `mar` | G |
| `apr` | H |
| `may` | I |
| `jun` | J |
| `jul` | K |
| `aug` | L |
| `sep` | M |
| `oct` | N |
| `nov` | O |
| `dec` | P |
| `peak_heating_1` | Q |
| `peak_heating_2` | R |
| `peak_cooling_1` | S |
| `peak_cooling_2` | T |
| `PER` | U |
| `latitude` | F |
| `longitude` | H |
| `elevation` | J |
| `elevation_unit` | M |
| `display_name` | L |
| `summer_delta_t` | N |
| `source` | P |

| Field | Row | Unit |
|-------|-----|------|
| `temperature_air` | 1 | C |
| `radiation_north` | 2 | KWH/M2 |
| `radiation_east` | 3 | KWH/M2 |
| `radiation_south` | 4 | KWH/M2 |
| `radiation_west` | 5 | KWH/M2 |
| `radiation_global` | 6 | KWH/M2 |
| `temperature_dewpoint` | 7 | C |
| `temperature_sky` | 8 | C |

### named_ranges

- `country` (named_range): `Klima_Region`
- `region` (named_range): `Klima_Region2`
- `data_set` (named_range): `Klima_Standort`

### defined_ranges

- `climate_zone` (address): `D13`
- `weather_station_altitude` (address): `D17`
- `site_altitude` (address): `D18`
- `latitude` (address): `F25`
- `longitude` (address): `H25`

---

## U-Values

Worksheet key: `UVALUES`

### constructor

- Header locator: col `L`, string `"Description of building assembly"`

| Field | Column | Unit |
|-------|--------|------|
| `display_name` | L |  |
| `r_si` | M | M2K/W |
| `r_se` | M | M2K/W |
| `interior_insulation` | R |  |
| `sec_1_description` | L |  |
| `sec_1_conductivity` | M | W/MK |
| `sec_2_description` | N |  |
| `sec_2_conductivity` | O | W/MK |
| `sec_3_description` | P |  |
| `sec_3_conductivity` | Q | W/MK |
| `thickness` | R | MM |
| `u_val_supplement` | R | W/M2K |
| `variants_layer_name` | E | |
| `variants_conductivity` | F | |
| `variants_thickness` | G | |
| `sec_2_percentage` | O |  |
| `sec_3_percentage` | Q |  |

**Row offsets within each assembly block:**

- `phpp_id_num_col_offset` (literal): `5`
- `name_row_offset` (literal): `2`
- `rsi_row_offset` (literal): `4`
- `rse_row_offset` (literal): `5`
- `first_layer_row_offset` (literal): `7`
- `last_layer_row_offset` (literal): `14`
- `result_val_row_offset` (literal): `19`
- `result_val_col` (literal): `R`
- `result_val_unit` (literal): `W/M2K`

---

## Areas

Worksheet key: `AREAS`

| Field Key | Label (`locator_string`) | Locator Col | Input Col | Row Offset | Unit | Options |
|-----------|-------------------------|-------------|-----------|------------|------|---------|
| `tfa_input` | 1-Treated floor area | M | T | 0 | M2 |  |

### summary_rows

- `temp_zones` (literal): `K`
- `area_type` (literal): `M`
- `group_number` (literal): `N`
- `area` (literal): `L`
- `average_u_value` (literal): `P`

### surface_rows

- Header locator: col `K`, string `"Area input"`
- Entry locator: col `K`, string `"1"`

| Field | Column | Unit |
|-------|--------|------|
| `description` | L |  |
| `group_number` | M |  |
| `quantity` | N |  |
| `area` | T | M2 |
| `assembly_id` | AA |  |
| `orientation` | AE |  |
| `angle` | AF |  |
| `shading` | AH |  |
| `absorptivity` | AI |  |
| `emissivity` | AJ |  |

### thermal_bridge_rows

- Header locator: col `K`, string `"Thermal bridge input"`
- Entry locator: col `K`, string `"1"`

| Field | Column | Unit |
|-------|--------|------|
| `description` | L |  |
| `group_number` | M |  |
| `quantity` | N |  |
| `length` | P | M |
| `psi_value` | V | W/MK |
| `fRsi_value` | X |  |

### defined_ranges

---

## Ground

Worksheet key: `GROUND`

---

## Components

Worksheet key: `COMPONENTS`

### glazings

- `entry_column` (literal): `IH`
- `entry_start_row` (literal): `13`
- `header_start_row` (literal): `8`

| Field | Column | Unit |
|-------|--------|------|
| `id` | IH |  |
| `description` | II |  |
| `g_value` | IJ |  |
| `u_value` | IK | W/M2K |

### frames

- Header locator: col `IO`, string `"Window and door frames"`
- Entry locator: col `IO`, string `"01ud"`

| Field | Column | Unit |
|-------|--------|------|
| `id` | IO |  |
| `description` | IP |  |
| `u_value_left` | IU | W/M2K |
| `u_value_right` | IW | W/M2K |
| `u_value_bottom` | JA | W/M2K |
| `u_value_top` | IY | W/M2K |
| `width_left` | IT | MM |
| `width_right` | IV | MM |
| `width_bottom` | IZ | MM |
| `width_top` | IX | MM |
| `psi_g_left` | IR | W/MK |
| `psi_g_right` | IR | W/MK |
| `psi_g_bottom` | IR | W/MK |
| `psi_g_top` | IR | W/MK |
| `psi_i_left` | JZ | W/MK |
| `psi_i_right` | JZ | W/MK |
| `psi_i_bottom` | KB | W/MK |
| `psi_i_top` | KA | W/MK |

### ventilators

- Header locator: col `LQ`, string `"Ventilation units"`
- Entry locator: col `LQ`, string `"01ud"`

| Field | Column | Unit |
|-------|--------|------|
| `id` | LQ |  |
| `display_name` | LR |  |
| `sensible_heat_recovery` | LS |  |
| `latent_heat_recovery` | LT |  |
| `electric_efficiency` | LW | WH/M3 |
| `min_m3h` | LX | M3/HR |
| `max_m3h` | LY | M3/HR |
| `pa_per_section` | LZ | PA |
| `pa_per_fittings` | MA | PA |
| `frost_protection_reqd` | MB |  |
| `noise_35DBA` | MC |  |
| `noise_supply_air` | MD |  |
| `noise_extract_air` | ME |  |
| `additional_info` | MF |  |

---

## Windows

Worksheet key: `WINDOWS`

### window_rows

- Header locator: col `L`, string `"Windows and entrance doors"`
- Entry locator: col `L`, string `"Quan-"`

| Field | Column | Unit |
|-------|--------|------|
| `quantity` | L |  |
| `description` | N |  |
| `orientation_angle` | O |  |
| `vertical_angle` | P |  |
| `orientation_label` | Q |  |
| `width` | R | M |
| `height` | S | M |
| `host` | T |  |
| `glazing_id` | U |  |
| `frame_id` | V |  |
| `psi_i_left` | AN | W/MK |
| `psi_i_right` | AO | W/MK |
| `psi_i_bottom` | AQ | W/MK |
| `psi_i_top` | AP | W/MK |
| `window_area` | AW | M2 |
| `glazing_area` | AX | M2 |
| `glazing_fraction` | AY |  |
| `u_w` | AZ | W/M2K |
| `u_w_installed` | BA | W/M2K |
| `comfort_exempt` | EQ |  |
| `comfort_temp` | EQ | C |
| `variant_input` | F |  |

### window_rows_end

- Header locator: col `L`, string `"Windows and entrance doors"`
- Entry locator: col `L`, string `"Unhide additional rows"`

---

## Shading

Worksheet key: `SHADING`

### shading_rows

- Header locator: col `T`, string `"Calculation of reduction factors for shading"`
- Entry locator: col `T`, string `"Quan-"`

| Field | Column | Unit |
|-------|--------|------|
| `h_hori` | AB | M |
| `d_hori` | AC | M |
| `o_reveal` | AD | M |
| `d_reveal` | AE | M |
| `o_over` | AF | M |
| `d_over` | AG | M |
| `r_other_winter` | AH |  |
| `r_other_summer` | AI |  |
| `temp_z` | AJ |  |
| `regulated` | AK |  |

### shading_rows_end

- Header locator: col `T`, string `"Shading"`
- Entry locator: col `T`, string `"Unhide additional rows"`

---

## Ventilation

Worksheet key: `VENTILATION`

| Field Key | Label (`locator_string`) | Locator Col | Input Col | Row Offset | Unit | Options |
|-----------|-------------------------|-------------|-----------|------------|------|---------|
| `vent_type` | Type of ventilation | I | K |  |  |  |
| `wind_coeff_e` | Wind protection coefficient, e | I | J |  |  |  |
| `wind_coeff_f` | Wind protection coefficient, f | I | J |  |  |  |
| `airtightness_n50` | Air change rate from pressurisation test | I | M |  |  |  |
| `airtightness_Vn50` | Net air volume for pressurisation test | I | M |  | M3 |  |
| `multi_unit_on` | 'Addl vent' worksheet | I | K |  |  |  |

**Configuration:**

- `variants_col` (literal): `D`

---

## Addl vent

Worksheet key: `ADDNL_VENT`

### rooms

- Header locator: col `C`, string `"Room"`
- Entry locator: col `C`, string `"1"`
- `last_col` (literal): `Z`

| Field | Column | Unit |
|-------|--------|------|
| `quantity` | D |  |
| `display_name` | E |  |
| `vent_unit_assigned` | F |  |
| `weighted_floor_area` | G | M2 |
| `clear_height` | H | M |
| `V_sup` | J | M3/HR |
| `V_eta` | K | M3/HR |
| `V_trans` | L | M3/HR |
| `operating_hours` | N |  |
| `operating_days` | O |  |
| `holiday_days` | P |  |
| `period_high_speed` | Q |  |
| `period_high_time` | R |  |
| `period_standard_speed` | S |  |
| `period_standard_time` | T |  |
| `period_minimum_speed` | U |  |
| `period_minimum_time` | V |  |

### units

- Header locator: col `C`, string `"Venti-"`
- Entry locator: col `C`, string `"1"`

| Field | Column | Unit |
|-------|--------|------|
| `quantity` | D |  |
| `display_name` | E |  |
| `unit_selected` | F |  |
| `oda_sup_pa` | K | PA |
| `eta_eha_pa` | L | PA |
| `addnl_pa` | M | PA |
| `ext_location` | Q |  |
| `subsoil_hr` | X |  |
| `frost_protection_type` | Z |  |
| `temperature_below_defrost_used` | AA | C |

### ducts

- Header locator: col `E`, string `"Round"`
- Entry locator: col `D`, string `"1"`

| Field | Column | Unit |
|-------|--------|------|
| `quantity` | D |  |
| `diameter` | F | MM |
| `width` | F | MM |
| `height` | G | MM |
| `insul_thickness` | H | MM |
| `insul_conductivity` | I | W/MK |
| `insul_reflective` | J |  |
| `sup_air_duct_len` | L | M |
| `oda_air_duct_len` | M | M |
| `exh_air_duct_len` | N | M |
| `duct_assign_1` | Q |  |
| `duct_assign_2` | R |  |
| `duct_assign_3` | S |  |
| `duct_assign_4` | T |  |
| `duct_assign_5` | U |  |
| `duct_assign_6` | V |  |
| `duct_assign_7` | W |  |
| `duct_assign_8` | X |  |
| `duct_assign_9` | Z |  |
| `duct_assign_10` | Z |  |

---

## Heating

Worksheet key: `HEATING_DEMAND`

**Configuration:**

- `unit` (literal): `kWh`
- `col_kWh_year` (literal): `O`
- `col_kWh_m2_year` (literal): `Q`
- `row_total_losses_transmission` (literal): `27`
- `row_total_losses_ventilation` (literal): `43`
- `row_total_losses` (literal): `45`
- `row_total_gains_solar` (literal): `61`
- `row_total_gains_internal` (literal): `65`
- `row_utilization_factor` (literal): `72`
- `row_useful_gains` (literal): `74`
- `row_annual_demand` (literal): `78`

---

## Heating load

Worksheet key: `HEATING_PEAK_LOAD`

**Configuration:**

- `unit` (literal): `W`
- `col_weather_1` (literal): `P`
- `col_weather_2` (literal): `R`
- `row_total_losses_transmission` (literal): `44`
- `row_total_losses_ventilation` (literal): `57`
- `row_total_losses` (literal): `60`
- `row_total_gains_solar` (literal): `73`
- `row_total_gains_internal` (literal): `77`
- `row_total_gains` (literal): `80`
- `row_total_load` (literal): `88`

---

## Cooling

Worksheet key: `COOLING_DEMAND`

**Configuration:**

- `unit` (literal): `kWh`
- `col_kWh_year` (literal): `O`
- `col_kWh_m2_year` (literal): `Q`
- `row_total_losses_transmission` (literal): `29`
- `row_total_losses_ventilation` (literal): `57`
- `row_total_losses` (literal): `59`
- `row_utilization_factor` (literal): `84`
- `row_useful_losses` (literal): `86`
- `row_total_gains_solar` (literal): `73`
- `row_total_gains_internal` (literal): `77`
- `row_total_gains` (literal): `79`
- `row_annual_sensible_demand` (literal): `88`
- `row_annual_latent_demand` (literal): `94`
- `address_specific_latent_cooling_demand` (address): `AN176`
- `address_tfa` (address): `O8`

---

## Cooling load

Worksheet key: `COOLING_PEAK_LOAD`

**Configuration:**

- `unit` (literal): `W`
- `col_weather_1` (literal): `P`
- `col_weather_2` (literal): `R`
- `row_total_losses_transmission` (literal): `35`
- `row_total_losses_ventilation` (literal): `43`
- `row_total_gains_solar` (literal): `56`
- `row_total_gains_internal` (literal): `60`
- `row_total_sensible_load` (literal): `64`
- `row_total_latent_load` (literal): `93`

---

## SummVent

Worksheet key: `SUMM_VENT`

---

## Cooling units

Worksheet key: `COOLING_UNITS`

**Configuration:**

- `SEER_unit` (literal): `W/W`

### supply_air

- `used` (named_range): `Kuehlgeraete_Zuluft_Kuehlung_Ankreuzen`
- `num_units` (named_range): `Kuehlgeraete_Kompressor_Zuluft_Anzahl`
- `device_type_name` (named_range): `Kuehlgeraete_Kompressor_Zuluft_Geraet`
- `SEER` (address): `X32`

### recirculation_air

- `used` (named_range): `Kuehlgeraete_Umluft_Kuehlung_Ankreuzen`
- `num_units` (named_range): `Kuehlgeraete_Kompressor_Umluft_Anzahl`
- `device_type_name` (named_range): `Kuehlgeraete_Kompressor_Umluft_Geraet`
- `SEER` (address): `X50`

### dehumidification

- `used` (named_range): `Kuehlgeraete_Zusaetzliche_Entfeuchtung_Ankreuzen`
- `waste_heat_to_room` (named_range): `Kuehlgeraete_Zusaetzliche_Entfeuchtung_Abwaerme`
- `SEER` (named_range): `Kuehlgeraete_Zusaetzliche_Entfeuchtung_JAZ`

### panel

- `used` (named_range): `Kuehlgeraete_Flaechenkuehlung_Ankreuzen`
- `device_type_name` (named_range): `Kuehlgeraete_Kompressor_Flaechenkuehlung_Geraet`
- `SEER` (address): `X69`

---

## DHW+Distribution

Worksheet key: `DHW`

### recirc_piping

- Header locator: col `D`, string `"DHW distribution"`
- Entry locator: col `E`, string `"DHW circulation pipes or, for heat interface units, forward and return flows"`
- `input_col_start` (literal): `J`

| Field | Row Offset | Unit |
|-------|-----------|------|
| `total_length` | 2 | M |
| `diameter` | 3 | MM |
| `insul_thickness` | 4 | MM |
| `insul_reflective` | 5 |  |
| `insul_conductivity` | 6 | W/MK |
| `daily_period` | 12 |  |
| `water_temp` | 13 | C |

### branch_piping

- Header locator: col `D`, string `"DHW distribution"`
- Entry locator: col `E`, string `"DHW stub pipes / individual pipes"`
- `input_col_start` (literal): `J`

| Field | Row Offset | Unit |
|-------|-----------|------|
| `water_temp` | 1 | C |
| `diameter` | 2 | MM |
| `total_length` | 3 | M |
| `num_taps` | 4 | |

### tanks

- Header locator: col `D`, string `"Storage heat losses"`
- Entry locator: col `J`, string `"Storage type 1"`
- `entry_row_start` (literal): `191`

| Field | Column |
|-------|--------|
| `tank_1` | J |
| `tank_2` | M |
| `tank_buffer` | P |

| Field | Row | Unit |
|-------|-----|------|
| `tank_type` | 0 | |
| `standby_losses` | 5 | W/K |
| `storage_capacity` | 6 | L |
| `standby_fraction` | 7 | |
| `tank_location` | 9 | |
| `water_temp` | 12 | C |

**tank_type options:** `0`: 0-No storage tank; `1`: 1-DHW and heating; `2`: 2-DHW only

**tank_location options:** `1`: 1-Inside; `0`: 2-Outside

---

## SolarDHW

Worksheet key: `SOLAR_DHW`

**Configuration:**

- `footprint_unit` (literal): `M2`
- `energy_unit` (literal): `KHW`

### ranges

- `footprint` (named_range): `SolarWW_Kollektorflaeche`
- `annual_dhw_contribution` (address): `N34`
- `annual_dhw_energy` (address): `P34`
- `annual_heating_contribution` (address): `N35`
- `annual_heating_energy` (address): `P35`

---

## PV

Worksheet key: `SOLAR_PV`

**Configuration:**

- `footprint_unit` (literal): `M2`
- `energy_unit` (literal): `KWH`

### columns

- `systems_start` (literal): `S`
- `systems_end` (literal): `W`

### rows

- `systems_start` (literal): `10`
- `current` (literal): `20`
- `voltage` (literal): `21`
- `num_panels` (literal): `29`
- `footprint` (literal): `37`
- `annual_energy` (literal): `42`
- `systems_end` (literal): `47`

---

## Electricity

Worksheet key: `ELECTRICITY`

### input_columns

- `selection` (literal): `E`
- `used` (literal): `E`
- `in_conditioned_space` (literal): `F`
- `energy_demand_per_use` (literal): `N`
- `utilization_factor` (literal): `H`
- `frequency` (literal): `K`
- `reference_quantity` (literal): `I`
- `annual_energy_demand` (literal): `AB`

### input_rows

| Appliance | Data Row | Selection Row | Options |
|-----------|----------|---------------|---------|
| `dishwasher` | 25 | 24 | `1`: 1-DHW connection; `2`: 2-Cold water connection |
| `clothes_washing` | 29 | 28 | `1`: 1-DHW connection; `2`: 2-Cold water connection |
| `clothes_drying` | 33 | 32 | `1`: 1-Clothes line; `2`: 2-Drying closet (cold!); `3`: 3-Drying closet (cold!) in extract air; `4`: 4-Condensation dryer; `5`: 5-Electric exhaust air dryer; `6`: 6-Gas exhaust air dryer |
| `refrigerator` | 16 |  |  |
| `freezer` | 17 |  |  |
| `fridge_freezer` | 18 |  |  |
| `cooking` | 21 | 20 | `1`: 1-Electricity; `2`: 2-Natural gas; `3`: 3-LPG |
| `lighting` | 37 |  |  |
| `lighting_interior` | 38 |  |  |
| `lighting_exterior` | 39 |  |  |
| `consumer_elec` | 68 |  |  |
| `small_appliances` | 68 |  |  |

---

## Use non-res

Worksheet key: `USE_NON_RES`

---

## Electricity non-res

Worksheet key: `ELEC_NON_RES`

### lighting_rows

- Header locator: col `C`, string `"Lighting"`
- Entry locator: col `C`, string `"Room / Zone"`

| Field | Column | Unit |
|-------|--------|------|
| `room_zone_name` | C | |
| `net_floor_area` | D | |
| `utilization_profile` | E | |
| `room_has_window` | F | |
| `room_angle_from_north` | G | |
| `room_orientation` | H | |
| `factor` | I | |
| `glazing_light_transmission` | J | |
| `room_depth` | K | |
| `room_width` | L | |
| `room_height` | M | |
| `lintel_height` | N | |
| `window_width` | O | |
| `daily_utilization` | P | |
| `nominal_illumination` | R | |
| `installed_power` | S | W/M2 |
| `lighting_control` | V | |
| `motion_detector_used` | W | |
| `utilization_hours_year` | X | |
| `ud_annual_full_load_hours` | Y | |
| `annual_full_load_hours` | Z | |
| `daily_full_load_hours` | AA | |
| `annual_energy_demand` | AC | |

---

## Aux Electricity

Worksheet key: `AUX_ELEC`

---

## IHG non-res

Worksheet key: `IHG_NON_RES`

---

## PER

Worksheet key: `PER`

**Configuration:**

- `locator_col` (literal): `P`
- `unit` (literal): `KWH`

### named_ranges

- `heating_type_1` (named_range): `PE_Waermeerzeuger_primaer`
- `heating_type_2` (named_range): `PE_Waermeerzeuger_sekundaer`

### columns

- `calculated_efficiency` (literal): `Q`
- `user_determined_efficiency` (literal): `R`
- `final_energy` (literal): `T`
- `per_energy` (literal): `V`
- `pe_energy` (literal): `X`
- `co2_emissions` (literal): `Z`

### addresses

- `tfa` (address): `Z7`
- `footprint` (address): `Z8`

### heating_types

- `range_start` (address): `P8`
- `range_end` (address): `T11`

### heating

- `locator_string_heading` (literal): `Heating`
- `locator_string_start` (literal): `Electricity (HP compact unit)`

### cooling

- `locator_string_heading` (literal): `Cooling and dehumidification`
- `locator_string_start` (literal): `Electricity cooling (HP)`

### dhw

- `locator_string_heading` (literal): `DHW generation`
- `locator_string_start` (literal): `Electricity (HP compact unit)`

### household_electric

- `locator_string_heading` (literal): `Occupant electricity + auxiliary electricity (other)`
- `locator_string_start` (literal): `User electricity (lighting, electrical devices, etc.)`

### additional_gas

- `locator_string_heading` (literal): `Additional gas demand`
- `locator_string_start` (literal): `Drying/Cooking`

### energy_generation

- `locator_string_heading` (literal): `Energy generation`
- `locator_string_start` (literal): `PV electricity`

---

## HP

Worksheet key: `HP`

---

## Boiler

Worksheet key: `BOILER`

---

## Data

Worksheet key: `DATA`

### version

- Header locator: col `A`, string `"PHPP Version"`
- Entry locator: col `B`, string `""`

---

## Overview

Worksheet key: `OVERVIEW`

### basic_data

- `address_number_dwellings_res` (address): `C27`
- `address_number_dwellings_nonres` (address): `E27`
- `address_number_occupants_res` (address): `C28`
- `address_number_occupants_nonres` (address): `E28`
- `address_project_name` (address): `C11`

### building_envelope

- `address_area_envelope`: col `C`, row `105` (M2)
- `address_area_tfa`: col `E`, row `105` (M2)

### ventilation

- `vn50`: col `C`, row `362` (M3)

---

## easyPH

Worksheet key: `EASY_PH`

---
