#!/usr/bin/env -S python3 -u
"""Post-Excel full-fidelity comparison (pure openpyxl, no programmatic Excel).

Usage:
    python scripts/verify_excel.py <source.xlsx> <written.xlsx>

Example workflow:
    # Step 1 — run the roundtrip to produce the written file:
    python scripts/roundtrip.py Data/Example.xlsx Data/Empty.xlsx
    #   → records/roundtrip_<stamp>/Example_written.xlsx

    # Step 2 — manually open BOTH files in Excel and save them:
    #   • Open Data/Example.xlsx in Excel, let it recalculate, Cmd-S, close.
    #   • Open records/.../Example_written.xlsx in Excel, Cmd-S, close.

    # Step 3 — run this script against the two Excel-saved files:
    python scripts/verify_excel.py Data/Example.xlsx \
        records/roundtrip_<stamp>/Example_written.xlsx

What this test does:
    After both files have been opened and saved in Excel, every formula
    cell has a fresh cached value. This script reads ALL mapped cells
    from both files via openpyxl (skip_formulas=False) and compares
    every field — inputs AND formula results. If every input was written
    correctly and Excel recalculated both files, all values should match.

    This is the definitive full-fidelity test. Unlike roundtrip.py
    Part 1 (which only verifies input cells) and Part 2 (which only
    checks cache freshness of the original), this test confirms that the
    written workbook produces the same calculation results as the original.

Results are saved to records/verify_excel_<timestamp>/:
    - <source>_all.json / <written>_all.json -- full field reads
    - verify_excel_summary.json -- counts only (checked/matched/mismatched)
    - verify_excel_details.txt -- every mismatch, untruncated (the terminal
      output above caps the printed list at 30)
"""

from __future__ import annotations

import json
import sys
import time
from datetime import datetime
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))

from phpp_tool.reader import read_phpp

FIELD_MAP = str(ROOT / "phpp-field-mapping" / "EN_10_6_IP.md")


def _count_values(data: dict, depth: int = 0) -> tuple[int, int]:
    """Count (total_values, non_none_values) recursively."""
    total = 0
    non_none = 0
    for v in data.values():
        if isinstance(v, dict):
            t, n = _count_values(v, depth + 1)
            total += t
            non_none += n
        elif isinstance(v, list):
            for item in v:
                if isinstance(item, dict):
                    t, n = _count_values(item, depth + 1)
                    total += t
                    non_none += n
                else:
                    total += 1
                    if item is not None:
                        non_none += 1
        else:
            total += 1
            if v is not None:
                non_none += 1
    return total, non_none


def _values_match(expected: Any, actual: Any) -> bool:
    """Compare values with tolerance for float rounding."""
    if expected == actual:
        return True
    if isinstance(expected, (int, float)) and isinstance(actual, (int, float)):
        if abs(expected) < 1e-10 and abs(actual) < 1e-10:
            return True
        if abs(expected) > 0:
            return abs(expected - actual) / abs(expected) < 1e-6
    if isinstance(expected, float) and isinstance(actual, int):
        return _values_match(expected, float(actual))
    if isinstance(expected, int) and isinstance(actual, float):
        return _values_match(float(expected), actual)
    return False


def _compare_deep(
    orig_data: dict[str, Any],
    written_data: dict[str, Any],
    path: str = "",
) -> tuple[int, int, int, list[str]]:
    """Recursively compare two nested reader dicts field by field.

    Returns (checked, matched, mismatched, details).
    Skips internal keys starting with '_' (e.g. _row, _config).
    """
    checked = 0
    matched = 0
    mismatched = 0
    details: list[str] = []

    all_keys = set(orig_data.keys()) | set(written_data.keys())
    for key in sorted(all_keys):
        if key.startswith("_"):
            continue
        full_key = f"{path}.{key}" if path else key
        orig_val = orig_data.get(key)
        written_val = written_data.get(key)

        if isinstance(orig_val, dict) and isinstance(written_val, dict):
            c, m, mm, d = _compare_deep(orig_val, written_val, full_key)
            checked += c
            matched += m
            mismatched += mm
            details.extend(d)
        elif isinstance(orig_val, list) and isinstance(written_val, list):
            max_len = max(len(orig_val), len(written_val))
            for i in range(max_len):
                ov = orig_val[i] if i < len(orig_val) else None
                wv = written_val[i] if i < len(written_val) else None
                item_key = f"{full_key}[{i}]"
                if isinstance(ov, dict) and isinstance(wv, dict):
                    c, m, mm, d = _compare_deep(ov, wv, item_key)
                    checked += c
                    matched += m
                    mismatched += mm
                    details.extend(d)
                else:
                    if ov is None and wv is None:
                        continue
                    checked += 1
                    if _values_match(ov, wv):
                        matched += 1
                    else:
                        mismatched += 1
                        details.append(
                            f"  {item_key}: original={ov!r}, written={wv!r}")
        else:
            if orig_val is None and written_val is None:
                continue
            checked += 1
            if _values_match(orig_val, written_val):
                matched += 1
            else:
                mismatched += 1
                details.append(
                    f"  {full_key}: original={orig_val!r},"
                    f" written={written_val!r}")

    return checked, matched, mismatched, details


def verify_excel(
    source_path: Path, written_path: Path, out_dir: Path,
) -> dict:
    """Compare ALL mapped cells between two Excel-saved files."""
    out_dir.mkdir(parents=True, exist_ok=True)

    print(f"\n{'=' * 60}")
    print(f"  POST-EXCEL VERIFICATION (pure openpyxl)")
    print(f"  Source:  {source_path.name}")
    print(f"  Written: {written_path.name}")
    print(f"{'=' * 60}")
    print()
    print("  Prerequisite: both files have been opened in Excel")
    print("  and saved, so all cached values are fresh.")
    print(f"  {'─' * 40}")

    # Step 1: read both files — all cells, no formula filtering
    print(f"\n[1/2] Reading {source_path.name} — all cells ...")
    t0 = time.time()
    orig_all = read_phpp(str(source_path), FIELD_MAP, skip_formulas=False)
    t_read_orig = time.time() - t0
    total_o, non_none_o = _count_values(orig_all)
    print(f"      {len(orig_all)} worksheets,"
          f" {non_none_o} non-None / {total_o} total values,"
          f" {t_read_orig:.1f}s")

    print(f"\n      Reading {written_path.name} — all cells ...")
    t0 = time.time()
    written_all = read_phpp(str(written_path), FIELD_MAP, skip_formulas=False)
    t_read_written = time.time() - t0
    total_w, non_none_w = _count_values(written_all)
    print(f"      {len(written_all)} worksheets,"
          f" {non_none_w} non-None / {total_w} total values,"
          f" {t_read_written:.1f}s")

    # Save full reads as JSON
    json_orig = out_dir / f"{source_path.stem}_all.json"
    json_orig.write_text(
        json.dumps(orig_all, indent=2, default=str), encoding="utf-8")
    json_written = out_dir / f"{written_path.stem}_all.json"
    json_written.write_text(
        json.dumps(written_all, indent=2, default=str), encoding="utf-8")

    # Step 2: deep compare — every mapped field
    print(f"\n[2/2] Comparing all mapped fields (inputs + formula results) ...")
    t0 = time.time()
    checked, matched, mismatched, details = _compare_deep(
        orig_all, written_all)
    t_compare = time.time() - t0

    if mismatched == 0:
        print(f"      *** All {checked} fields match —"
              f" FULL FIDELITY CONFIRMED ***")
    else:
        print(f"      {checked} checked, {matched} matched,"
              f" {mismatched} MISMATCHES:")
        for d in details[:30]:
            print(d)
        if len(details) > 30:
            print(f"      ... and {len(details) - 30} more"
                  f" (full list saved to verify_excel_details.txt)")

    total_t = t_read_orig + t_read_written + t_compare
    print(f"\n      {t_compare:.1f}s compare,"
          f" {total_t:.1f}s total")

    print(f"\n  {'═' * 40}")
    print(f"  Artifacts saved to {out_dir}/")

    result = {
        "source": source_path.name,
        "written": written_path.name,
        "orig_non_none": non_none_o,
        "written_non_none": non_none_w,
        "fields_checked": checked,
        "fields_matched": matched,
        "fields_mismatched": mismatched,
        "time_total": total_t,
    }

    summary_path = out_dir / "verify_excel_summary.json"
    summary_path.write_text(
        json.dumps(result, indent=2, default=str), encoding="utf-8")

    # Full, untruncated mismatch details -- the screen output above caps at
    # 30 for readability, but this file keeps every one so a later session
    # doesn't have to reconstruct them from the *_all.json dumps by hand.
    details_path = out_dir / "verify_excel_details.txt"
    if mismatched == 0:
        details_text = f"All {checked} fields match -- FULL FIDELITY CONFIRMED\n"
    else:
        details_text = (
            f"{checked} checked, {matched} matched, {mismatched} MISMATCHES:\n"
            + "\n".join(details) + "\n"
        )
    details_path.write_text(details_text, encoding="utf-8")

    return result


def main() -> None:
    if len(sys.argv) != 3:
        print(__doc__)
        sys.exit(1)

    source_path = Path(sys.argv[1])
    written_path = Path(sys.argv[2])

    if not source_path.exists():
        print(f"Source file not found: {source_path}")
        sys.exit(1)
    if not written_path.exists():
        print(f"Written file not found: {written_path}")
        sys.exit(1)

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_dir = ROOT / "records" / f"verify_excel_{stamp}"

    verify_excel(source_path, written_path, out_dir)


if __name__ == "__main__":
    main()
