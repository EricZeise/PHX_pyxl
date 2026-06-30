"""CLI for the PHPP Data Tool (openpyxl backend).

Commands:
    phpp-tool read  <filled.xlsx> -o record.json
    phpp-tool write <record.json> <blank.xlsx> -o populated.xlsx
    phpp-tool inspect-map
"""

from __future__ import annotations

from pathlib import Path

import click

from phpp_tool.map_parser import parse_field_map

DEFAULT_FIELD_MAP = "phpp-field-mapping.md"


@click.group()
@click.version_option(package_name="phpp-data-tool-pyxl")
def main() -> None:
    """PHPP Data Tool — read and write Passive House input data (openpyxl)."""


@main.command()
@click.argument("workbook", type=click.Path(exists=True, dir_okay=False))
@click.option("-o", "--output", type=click.Path(dir_okay=False),
              help="Output JSON file path.")
@click.option(
    "--field-map",
    type=click.Path(exists=True, dir_okay=False),
    default=DEFAULT_FIELD_MAP,
    show_default=True,
    help="Path to the field map markdown file.",
)
def read(workbook: str, output: str | None, field_map: str) -> None:
    """Read a filled PHPP workbook into a JSON building record."""
    from phpp_tool.models import BuildingRecord
    from phpp_tool.reader import read_phpp

    data = read_phpp(workbook, field_map)
    record = BuildingRecord.from_reader_dict(data)
    json_str = record.to_json()

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
@click.option(
    "--field-map",
    type=click.Path(exists=True, dir_okay=False),
    default=DEFAULT_FIELD_MAP,
    show_default=True,
    help="Path to the field map markdown file.",
)
def write(
        record_file: str,
        template: str,
        output: str,
        field_map: str) -> None:
    """Write a JSON building record into a blank PHPP workbook."""
    from phpp_tool.models import BuildingRecord
    from phpp_tool.writer import write_phpp

    raw = Path(record_file).read_text(encoding="utf-8")
    record = BuildingRecord.model_validate_json(raw)
    record_dict = record.model_dump(exclude_none=True)

    Path(output).parent.mkdir(parents=True, exist_ok=True)
    write_phpp(record_dict, template, output, field_map)
    click.echo(f"Written to {output}")


@main.command("inspect-map")
@click.option(
    "--field-map",
    type=click.Path(exists=True, dir_okay=False),
    default=DEFAULT_FIELD_MAP,
    show_default=True,
    help="Path to the field map markdown file.",
)
def inspect_map(field_map: str) -> None:
    """Inspect the field map — list all mapped sheets and field counts."""
    fm = parse_field_map(field_map)

    click.echo(f"Field map: {field_map}")
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
