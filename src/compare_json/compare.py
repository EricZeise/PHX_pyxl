"""Deep comparison of two JSON files with categorized difference reporting."""

from __future__ import annotations

import json
from collections import Counter
from typing import Any


FLOAT_TOLERANCE = 1e-10


def count_fields(d: dict | list) -> int:
    """Count total leaf fields in a nested structure."""
    n = 0
    if isinstance(d, dict):
        for v in d.values():
            if isinstance(v, (dict, list)):
                n += count_fields(v)
            else:
                n += 1
    elif isinstance(d, list):
        for item in d:
            if isinstance(item, (dict, list)):
                n += count_fields(item)
            else:
                n += 1
    return n


def diff_dicts(
    d1: dict, d2: dict, path: str = "",
) -> list[dict[str, Any]]:
    """Compare two dicts, returning difference records."""
    diffs: list[dict[str, Any]] = []
    all_keys = sorted(d1.keys() | d2.keys())

    for k in all_keys:
        p = f"{path}.{k}" if path else k

        if k not in d1:
            diffs.append({"type": "added", "path": p, "val": d2[k]})
        elif k not in d2:
            diffs.append({"type": "missing", "path": p, "val": d1[k]})
        elif isinstance(d1[k], dict) and isinstance(d2[k], dict):
            diffs.extend(diff_dicts(d1[k], d2[k], p))
        elif isinstance(d1[k], list) and isinstance(d2[k], list):
            diffs.extend(_diff_lists(d1[k], d2[k], p))
        elif d1[k] != d2[k]:
            diffs.append(_classify_value_diff(p, d1[k], d2[k]))

    return diffs


def _diff_lists(l1: list, l2: list, path: str) -> list[dict[str, Any]]:
    """Compare two lists element-by-element."""
    diffs: list[dict[str, Any]] = []

    if len(l1) != len(l2):
        diffs.append({
            "type": "list_len",
            "path": path,
            "orig": len(l1),
            "rt": len(l2),
        })

    for i in range(min(len(l1), len(l2))):
        ip = f"{path}[{i}]"
        if isinstance(l1[i], dict) and isinstance(l2[i], dict):
            diffs.extend(diff_dicts(l1[i], l2[i], ip))
        elif l1[i] != l2[i]:
            diffs.append(_classify_value_diff(ip, l1[i], l2[i]))

    return diffs


def _classify_value_diff(path: str, orig: Any, rt: Any) -> dict[str, Any]:
    """Classify a single value difference."""
    if isinstance(orig, (int, float)) and isinstance(rt, (int, float)):
        if abs(orig - rt) < FLOAT_TOLERANCE:
            return {
                "type": "float_rounding",
                "path": path,
                "orig": orig,
                "rt": rt}
        return {"type": "value_diff", "path": path, "orig": orig, "rt": rt}

    if orig is None and rt is not None:
        return {"type": "null_to_value", "path": path, "rt": rt}
    if orig is not None and rt is None:
        return {"type": "value_to_null", "path": path, "orig": orig}

    return {"type": "value_diff", "path": path, "orig": orig, "rt": rt}


def compare(file1: dict, file2: dict) -> dict[str, Any]:
    """Compare two parsed JSON structures and return a results dict."""
    diffs = diff_dicts(file1, file2)
    total_f1 = count_fields(file1)
    total_f2 = count_fields(file2)
    cats = Counter(d["type"] for d in diffs)

    # Every diff belongs to exactly one bucket below (value_diff further
    # splits on whether it's just _row metadata drift) -- one pass instead
    # of filtering the same list eight separate times.
    row_diffs: list[dict[str, Any]] = []
    real_value: list[dict[str, Any]] = []
    v2n: list[dict[str, Any]] = []
    n2v: list[dict[str, Any]] = []
    missing: list[dict[str, Any]] = []
    added: list[dict[str, Any]] = []
    fr: list[dict[str, Any]] = []
    ll: list[dict[str, Any]] = []

    for d in diffs:
        t = d["type"]
        if t == "value_diff":
            (row_diffs if "_row" in d["path"] else real_value).append(d)
        elif t == "value_to_null":
            v2n.append(d)
        elif t == "null_to_value":
            n2v.append(d)
        elif t == "missing":
            missing.append(d)
        elif t == "added":
            added.append(d)
        elif t == "float_rounding":
            fr.append(d)
        elif t == "list_len":
            ll.append(d)

    meaningful = len(real_value) + len(v2n) + \
        len(n2v) + len(missing) + len(added)
    cosmetic = len(fr) + len(row_diffs)
    structural = len(ll)

    return {
        "total_fields_file1": total_f1,
        "total_fields_file2": total_f2,
        "total_diffs": len(diffs),
        "categories": dict(cats.most_common()),
        "meaningful": meaningful,
        "cosmetic": cosmetic,
        "structural": structural,
        "data_fidelity_pct": (
            (1 - meaningful / total_f1) * 100
            if total_f1 else 100.0
        ),
        "match_rate_pct": (
            (1 - len(diffs) / total_f1) * 100
            if total_f1 else 100.0
        ),
        "max_float_error": max(
            (abs(d["orig"] - d["rt"]) for d in fr),
            default=0.0,
        ),
        "diffs": diffs,
        "row_diffs": row_diffs,
        "real_value_diffs": real_value,
        "values_lost": v2n,
        "values_gained": n2v,
        "missing_keys": missing,
        "added_keys": added,
        "float_rounding": fr,
        "list_length": ll,
    }


def format_report(
    results: dict[str, Any],
    file1_name: str = "File 1",
    file2_name: str = "File 2",
) -> str:
    """Format comparison results as a human-readable report string."""
    lines: list[str] = []

    lines.append("=" * 70)
    lines.append("  JSON COMPARISON REPORT")
    lines.append("=" * 70)
    lines.append("")
    lines.append(f"  File 1: {file1_name}")
    lines.append(f"  File 2: {file2_name}")
    lines.append("")
    lines.append(f"  File 1 fields:  {results['total_fields_file1']:,}")
    lines.append(f"  File 2 fields:  {results['total_fields_file2']:,}")
    lines.append(f"  Total differences:   {results['total_diffs']:,}")
    lines.append(f"  Match rate:          {results['match_rate_pct']:.1f}%")
    lines.append("")

    # Category breakdown
    lines.append("-" * 70)
    lines.append("  CATEGORY BREAKDOWN")
    lines.append("-" * 70)
    lines.append("")

    cat_labels = {
        "float_rounding": "Float rounding (< 1e-10, insignificant)",
        "value_diff": "Value differences (data mismatch)",
        "value_to_null": "Value in File 1 → None in File 2 (lost)",
        "null_to_value": "None in File 1 → value in File 2 (gained)",
        "missing": "Keys in File 1 but missing in File 2",
        "added": "Keys in File 2 but not in File 1",
        "list_len": "List length differences",
    }
    total = results["total_diffs"] or 1
    for cat, count in results["categories"].items():
        label = cat_labels.get(cat, cat)
        pct = count / total * 100
        lines.append(f"  {label:50s} {count:>5}  ({pct:5.1f}%)")
    lines.append("")

    # Value differences detail
    row_diffs = results["row_diffs"]
    real_value = results["real_value_diffs"]

    lines.append("-" * 70)
    lines.append("  VALUE DIFFERENCES DETAIL")
    lines.append("-" * 70)
    lines.append("")
    lines.append(
        f"  Row position shifts (_row metadata):    {len(row_diffs):>5}")
    lines.append(
        f"  Actual data value mismatches:            {len(real_value):>5}")
    lines.append("")

    if real_value:
        by_ws = _group_by_top_key(real_value)
        lines.append("  By section:")
        for ws in sorted(by_ws.keys()):
            lines.append(f"    {ws}: {len(by_ws[ws])} mismatches")
        lines.append("")

        for ws in sorted(by_ws.keys()):
            lines.append(f"  --- {ws} ---")
            for d in by_ws[ws]:
                short = d["path"][len(ws) + 1:]
                ov = json.dumps(d.get("orig"))[:45]
                rv = json.dumps(d.get("rt"))[:45]
                lines.append(f"    {short:50s} {ov:>45s}  →  {rv}")
            lines.append("")

    # Values lost
    v2n = results["values_lost"]
    if v2n:
        lines.append("-" * 70)
        lines.append(
            f"  VALUES LOST (File 1 had value, File 2 has None): {len(v2n)}")
        lines.append("-" * 70)
        lines.append("")
        by_ws = _group_by_top_key(v2n)
        for ws in sorted(by_ws.keys()):
            lines.append(f"  --- {ws}: {len(by_ws[ws])} lost ---")
            for d in by_ws[ws][:10]:
                short = d["path"][len(ws) + 1:]
                lines.append(
                    f"    {short:50s} was: {json.dumps(d['orig'])[:50]}")
            if len(by_ws[ws]) > 10:
                lines.append(f"    ... and {len(by_ws[ws]) - 10} more")
            lines.append("")

    # Values gained
    n2v = results["values_gained"]
    if n2v:
        lines.append("-" * 70)
        lines.append(
            f"  VALUES GAINED (File 1 had None, File 2 has value): {len(n2v)}")
        lines.append("-" * 70)
        lines.append("")
        by_ws = _group_by_top_key(n2v)
        for ws in sorted(by_ws.keys()):
            lines.append(f"  --- {ws}: {len(by_ws[ws])} gained ---")
            for d in by_ws[ws][:10]:
                short = d["path"][len(ws) + 1:]
                lines.append(
                    f"    {short:50s} now: {json.dumps(d['rt'])[:50]}")
            if len(by_ws[ws]) > 10:
                lines.append(f"    ... and {len(by_ws[ws]) - 10} more")
            lines.append("")

    # List length diffs
    ll = results["list_length"]
    if ll:
        lines.append("-" * 70)
        lines.append("  LIST LENGTH DIFFERENCES")
        lines.append("-" * 70)
        lines.append("")
        for d in ll:
            delta = d["rt"] - d["orig"]
            lines.append(
                f"  {d['path']:50s} {d['orig']:>4} rows"
                f" → {d['rt']:>4} rows  (delta: {delta:+d})")
        lines.append("")

    # Missing keys
    mk = results["missing_keys"]
    if mk:
        lines.append("-" * 70)
        lines.append(f"  KEYS IN FILE 1 BUT MISSING IN FILE 2: {len(mk)}")
        lines.append("-" * 70)
        lines.append("")
        for d in mk[:30]:
            lines.append(f"  {d['path']}")
        if len(mk) > 30:
            lines.append(f"  ... and {len(mk) - 30} more")
        lines.append("")

    # Added keys
    ak = results["added_keys"]
    if ak:
        lines.append("-" * 70)
        lines.append(f"  KEYS IN FILE 2 BUT NOT IN FILE 1: {len(ak)}")
        lines.append("-" * 70)
        lines.append("")
        for d in ak[:30]:
            lines.append(f"  {d['path']}")
        if len(ak) > 30:
            lines.append(f"  ... and {len(ak) - 30} more")
        lines.append("")

    # Float rounding
    fr = results["float_rounding"]
    if fr:
        lines.append("-" * 70)
        lines.append(f"  FLOAT ROUNDING: {len(fr)} instances")
        lines.append(f"  Maximum error: {results['max_float_error']:.2e}")
        lines.append(
            "  These are insignificant floating-point"
            " representation differences.")
        lines.append("-" * 70)
        lines.append("")

    # Summary
    lines.append("=" * 70)
    lines.append("  SUMMARY")
    lines.append("=" * 70)
    lines.append("")
    lines.append(
        f"  Meaningful differences:    {results['meaningful']:>5}"
        f"  (data mismatches, lost/gained values, missing keys)"
    )
    lines.append(
        f"  Cosmetic differences:      {results['cosmetic']:>5}"
        f"  (float rounding + row position metadata)"
    )
    lines.append(
        f"  Structural differences:    {results['structural']:>5}"
        f"  (list length mismatches)"
    )
    lines.append(f"  Total:                     {results['total_diffs']:>5}")
    lines.append("")
    lines.append(
        f"  Data fidelity:             {results['data_fidelity_pct']:.1f}%"
        f"  (excluding cosmetic/structural)"
    )
    lines.append("")

    if results["total_diffs"] == 0:
        lines.append("  *** PERFECT MATCH — no differences found ***")
        lines.append("")

    return "\n".join(lines)


def _group_by_top_key(diffs: list[dict]) -> dict[str, list[dict]]:
    """Group diffs by their top-level key (first path segment)."""
    by_key: dict[str, list[dict]] = {}
    for d in diffs:
        key = d["path"].split(".")[0]
        by_key.setdefault(key, []).append(d)
    return by_key
