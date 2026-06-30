"""CLI for comparing two JSON files."""

from __future__ import annotations

import json
from pathlib import Path

import click

from compare_json.compare import compare, format_report


@click.command()
@click.argument("file1", type=click.Path(exists=True, dir_okay=False))
@click.argument("file2", type=click.Path(exists=True, dir_okay=False))
@click.option(
    "-o", "--output",
    type=click.Path(dir_okay=False),
    help="Save the report to a file.",
)
@click.option(
    "--json-output",
    type=click.Path(dir_okay=False),
    help="Save raw comparison data as JSON.",
)
def main(
        file1: str,
        file2: str,
        output: str | None,
        json_output: str | None) -> None:
    """Compare two JSON files and report differences.

    FILE1 is the reference (original) JSON file.
    FILE2 is the JSON file to compare against.
    """
    with open(file1, encoding="utf-8") as f:
        data1 = json.load(f)
    with open(file2, encoding="utf-8") as f:
        data2 = json.load(f)

    results = compare(data1, data2)
    report = format_report(results, file1_name=file1, file2_name=file2)

    click.echo(report)

    if output:
        Path(output).parent.mkdir(parents=True, exist_ok=True)
        Path(output).write_text(report + "\n", encoding="utf-8")
        click.echo(f"\nReport saved to: {output}")

    if json_output:
        Path(json_output).parent.mkdir(parents=True, exist_ok=True)
        export = {
            "file1": file1,
            "file2": file2,
            "total_fields_file1": results["total_fields_file1"],
            "total_fields_file2": results["total_fields_file2"],
            "total_diffs": results["total_diffs"],
            "match_rate_pct": results["match_rate_pct"],
            "data_fidelity_pct": results["data_fidelity_pct"],
            "meaningful": results["meaningful"],
            "cosmetic": results["cosmetic"],
            "structural": results["structural"],
            "categories": results["categories"],
            "max_float_error": results["max_float_error"],
        }
        Path(json_output).write_text(
            json.dumps(export, indent=2), encoding="utf-8"
        )
        click.echo(f"JSON results saved to: {json_output}")
