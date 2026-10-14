# -*- coding: utf-8 -*-
"""Compare a newly reproduced output directory against the archived evidence files."""
from __future__ import annotations
import argparse
import hashlib
from pathlib import Path

FILES = [
    "abm_convergence_raw.csv",
    "abm_convergence_summary.csv",
    "abm_scenario_raw.csv",
    "abm_scenario_summary.csv",
    "abm_delay_intervention_raw.csv",
    "abm_delay_intervention_summary.csv",
    "abm_oat_robustness.csv",
    "abm_model_metadata.json",
    "abm_scenarios.json",
]

def digest(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for block in iter(lambda: f.read(1024 * 1024), b""):
            h.update(block)
    return h.hexdigest()

def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--expected", required=True)
    ap.add_argument("--actual", required=True)
    args = ap.parse_args()
    expected = Path(args.expected)
    actual = Path(args.actual)
    failed = False
    for name in FILES:
        e = expected / name
        a = actual / name
        if not e.exists() or not a.exists():
            print(f"MISSING  {name}")
            failed = True
            continue
        ok = digest(e) == digest(a)
        print(f"{'MATCH' if ok else 'DIFFER'}  {name}")
        failed = failed or not ok
    raise SystemExit(1 if failed else 0)

if __name__ == "__main__":
    main()
