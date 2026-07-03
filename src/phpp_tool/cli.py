"""CLI for the PHPP Data Tool (openpyxl backend).

Commands:
    phpp-tool read  <filled.xlsx> -o record.json --phpp-version EN_10_6_IP
    phpp-tool write <record.json> <blank.xlsx> -o populated.xlsx --phpp-version EN_10_6_IP
    phpp-tool inspect-map --phpp-version EN_10_6_SI

--phpp-version selects a versioned field map from phpp-field-mapping/<version>.md
-- e.g. EN_10_6_IP (IP-shell workbook with SI mirror tabs) or EN_10_6_SI
(genuinely SI-native, single-shell workbook). read and write must use the
same version/type for a given workbook -- write warns if the record's
stamped _phpp_version doesn't match. --field-map overrides with a direct
path to a specific field map file, bypassing version resolution.
"""

from __future__ import annotations

import json
from pathlib import Path

import click

from phpp_tool.map_parser import parse_field_map

FIELD_MAP_DIR = "phpp-field-mapping"
DEFAULT_PHPP_VERSION = "EN_10_6_IP"


def _resolve_field_map(phpp_version: str, field_map: str | None) -> str:
    """Return an explicit --field-map path if given, else resolve by version."""
    if field_map:
        return field_map
    path = Path(FIELD_MAP_DIR) / f"{phpp_version}.md"
    if not path.exists():
        raise click.ClickException(
            f"No field map found for --phpp-version {phpp_version!r} "
            f"(expected {path})")
    return str(path)


_phpp_version_option = click.option(
    "--phpp-version",
    default=DEFAULT_PHPP_VERSION,
    show_default=True,
    help=f"PHPP version/variant, resolved to {FIELD_MAP_DIR}/<version>.md.",
)
_field_map_option = click.option(
    "--field-map",
    type=click.Path(exists=True, dir_okay=False),
    default=None,
    help="Path to a specific field map file, overriding --phpp-version.",
)


@click.group()
@click.version_option(package_name="phpp-data-tool-pyxl")
def main() -> None:
    """PHPP Data Tool — read and write Passive House input data (openpyxl)."""


@main.command()
@click.argument("workbook", type=click.Path(exists=True, dir_okay=False))
@click.option("-o", "--output", type=click.Path(dir_okay=False),
              help="Output JSON file path.")
@_phpp_version_option
@_field_map_option
def read(workbook: str, output: str | None, phpp_version: str,
         field_map: str | None) -> None:
    """Read a filled PHPP workbook into a JSON building record."""
    from phpp_tool.models import BuildingRecord
    from phpp_tool.reader import read_phpp

    resolved_map = _resolve_field_map(phpp_version, field_map)
    data = read_phpp(workbook, resolved_map)
    record = BuildingRecord.from_reader_dict(data)
    record_dict = record.model_dump(mode="json", exclude_none=True)
    record_dict["_phpp_version"] = phpp_version
    json_str = json.dumps(record_dict, indent=2)

    if output:
        Path(output).parent.mkdir(parents=True, exist_ok=True)
        Path(output).write_text(json_str, encoding="utf-8")
        click.echo(f"Written to {output}")
    else:
        click.echo(json_str)


@main.command()
@click.argument("record_file", type=click.Path(exists=True, dir_okay=False))
@click.argument("template", type=click.Path(exists=True, dir_okay=False))
@click.option("-o", "--output", type=click.Path(dir_okay=False),
              required=True, help="Output PHPP workbook path.")
@_phpp_version_option
@_field_map_option
def write(
        record_file: str,
        template: str,
        output: str,
        phpp_version: str,
        field_map: str | None) -> None:
    """Write a JSON building record into a blank PHPP workbook."""
    from phpp_tool.models import BuildingRecord
    from phpp_tool.writer import write_phpp

    resolved_map = _resolve_field_map(phpp_version, field_map)
    raw = Path(record_file).read_text(encoding="utf-8")
    raw_dict = json.loads(raw)
    source_version = raw_dict.get("_phpp_version")
    if source_version and source_version != phpp_version:
        click.echo(
            f"WARNING: record was read with --phpp-version "
            f"{source_version!r} but this write uses {phpp_version!r} -- "
            f"the record and template should be the same PHPP "
            f"version/type.", err=True)

    record = BuildingRecord.model_validate_json(raw)
    record_dict = record.model_dump(exclude_none=True)

    Path(output).parent.mkdir(parents=True, exist_ok=True)
    write_phpp(record_dict, template, output, resolved_map)
    click.echo(f"Written to {output}")


@main.command("inspect-map")
@_phpp_version_option
@_field_map_option
def inspect_map(phpp_version: str, field_map: str | None) -> None:
    """Inspect the field map — list all mapped sheets and field counts."""
    resolved_map = _resolve_field_map(phpp_version, field_map)
    fm = parse_field_map(resolved_map)

    click.echo(f"Field map: {resolved_map}")
    click.echo(f"Worksheets mapped: {len(fm)}\n")

    for ws_key, ws_spec in fm.items():
        sheet_name = ws_spec.get("sheet_name", ws_key)
        field_count = len(ws_spec.get("fields", {}))
        section_count = len(ws_spec.get("sections", {}))
        config_count = len(ws_spec.get("config", {}))

        parts = []
        if field_count:
            parts.append(f"{field_count} fields")
        if section_count:
            parts.append(f"{section_count} sections")
        if config_count:
            parts.append(f"{config_count} config")

        detail = ", ".join(parts) if parts else "(stub)"
        click.echo(f"  {ws_key:25s} ({sheet_name}) — {detail}")
