"""The Step 2 ratchet (CodeReview.md): per-rule counts of the report-only Ruff families and of the suppression
comments may only go down.

    python scripts/check_ratchet.py            # compare with scripts/ratchet_baseline.json; exit 1 if anything rose
    python scripts/check_ratchet.py --update   # write the current counts as the new baseline (after fixing some)

Tracked: Ruff's `C90` (complexity) and `PLR` (design limits) rules, which are not enforced in `pyproject.toml`, and the
number of `# noqa` and `# ty: ignore` comments in the application, tests and tools.
"""

import collections
import json
import re
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
BASELINE = Path(__file__).resolve().parent / "ratchet_baseline.json"
REPORT_ONLY_RULES = "C90,PLR"
SUPPRESSION_DIRS = ("src", "tests", "scripts")
SUPPRESSIONS = {
    "noqa": re.compile(r"#\s*noqa\b"),
    "ty_ignore": re.compile(r"#\s*ty:\s*ignore\b"),
}


def ruff_counts() -> dict[str, int]:
    run = subprocess.run(
        [
            sys.executable,
            "-m",
            "ruff",
            "check",
            ".",
            "--select",
            REPORT_ONLY_RULES,
            "--output-format",
            "json",
            "--exit-zero",
        ],
        capture_output=True,
        text=True,
        cwd=ROOT,
        check=True,
    )
    return dict(collections.Counter(item["code"] for item in json.loads(run.stdout)))


def suppression_counts() -> dict[str, int]:
    counts = dict.fromkeys(SUPPRESSIONS, 0)
    for folder in SUPPRESSION_DIRS:
        for path in (ROOT / folder).rglob("*.py"):
            if path.name == Path(__file__).name or any(
                part.startswith("2026-10-31") for part in path.parts
            ):
                continue  # this file (it names the comments) and the external review's evidence scripts
            for line in path.read_text().splitlines():
                for name, pattern in SUPPRESSIONS.items():
                    if pattern.search(line):
                        counts[name] += 1
    return counts


def current() -> dict[str, dict[str, int]]:
    return {
        "ruff": dict(sorted(ruff_counts().items())),
        "suppressions": suppression_counts(),
    }


def main(argv: list[str]) -> int:
    now = current()
    if "--update" in argv:
        BASELINE.write_text(json.dumps(now, indent=2) + "\n")
        print(
            f"wrote {BASELINE.name}: {sum(now['ruff'].values())} Ruff findings, suppressions {now['suppressions']}"
        )
        return 0
    before = json.loads(BASELINE.read_text())
    rose, fell = [], []
    for group in ("ruff", "suppressions"):
        for key in sorted(set(before[group]) | set(now[group])):
            old, new = before[group].get(key, 0), now[group].get(key, 0)
            if new > old:
                rose.append(f"{key}: {old} -> {new}")
            elif new < old:
                fell.append(f"{key}: {old} -> {new}")
    for line in fell:
        print(f"down  {line}")
    for line in rose:
        print(f"UP    {line}")
    if rose:
        print(
            "FAIL  a tracked count rose; fix the new findings (or justify them and run with --update)"
        )
        return 1
    if fell:
        print("lower the baseline: python scripts/check_ratchet.py --update")
    print("ratchet ok")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
