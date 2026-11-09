"""Runs the regression suite under coverage (branch coverage of the packages under `src/`, child processes included) and
compares the total with scripts/coverage_baseline.json (CodeReview.md, Step 2).

    python scripts/run_with_coverage.py            # run, report, fail if coverage dropped by more than 0.5 points
    python scripts/run_with_coverage.py --update   # record the current total as the baseline

Writes coverage.json (and prints the per-file table) in the repository root; both are git-ignored build output.
"""

import json
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
BASELINE = Path(__file__).resolve().parent / "coverage_baseline.json"
TOLERANCE_POINTS = 0.5


def run(*args: str) -> subprocess.CompletedProcess:
    return subprocess.run(
        [sys.executable, "-m", "coverage", *args], cwd=ROOT, check=False
    )


def main(argv: list[str]) -> int:
    run("erase")
    suite = run("run", "-m", "pytest", "-q", "tests")
    if suite.returncode != 0:
        print("FAIL  the suite failed under coverage")
        return 1
    run("combine", "--quiet")
    run("json", "-o", "coverage.json", "--quiet")
    run("report", "--skip-covered", "--sort=cover")
    totals = json.loads((ROOT / "coverage.json").read_text())["totals"]
    percent = round(totals["percent_covered"], 2)
    print(f"\ntotal (statements + branches): {percent} %")
    if "--update" in argv:
        BASELINE.write_text(json.dumps({"percent_covered": percent}, indent=2) + "\n")
        print(f"wrote {BASELINE.name}")
        return 0
    baseline = json.loads(BASELINE.read_text())["percent_covered"]
    if percent < baseline - TOLERANCE_POINTS:
        print(f"FAIL  coverage fell from {baseline} % to {percent} %")
        return 1
    print(f"coverage ok (baseline {baseline} %)")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
