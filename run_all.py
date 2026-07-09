"""
run_all.py - one-click ECL pipeline.
Runs every stage in order, stops on the first failure, prints a timing summary.

Usage:  python run_all.py
Output: ecl.db, intermediate CSVs, ECL_Report.xlsx, validation_report.xlsx
"""

import subprocess
import sys
import time

PHASES = [
    ("Phase 0  data + SQLite",         "base_loans.py"),
    ("Phase 1  SQL refactor",          "sql_refactor.py"),
    ("Phase 3  chain ladder",          "chain_ladder.py"),
    ("Phase 4  movements + loss rate", "loss_rate.py"),
    ("Phase 5  final ECL",             "final_ecl.py"),
    ("Phase 6  Excel report",          "report.py"),
    ("Phase 7  validation",            "validation.py"),
]

def main():
    print("=" * 64)
    print("ECL AUTOMATION - full pipeline")
    print("=" * 64)
    t0 = time.time()
    timings = []
    for label, script in PHASES:
        print(f"\n>>> {label}  ({script})")
        t = time.time()
        r = subprocess.run([sys.executable, script], capture_output=True, text=True)
        dt = time.time() - t
        timings.append((label, dt, r.returncode))
        # echo the last few lines of each phase's output
        tail = [ln for ln in r.stdout.strip().splitlines() if ln][-4:]
        for ln in tail:
            print("    " + ln)
        if r.returncode != 0:
            print(f"\n!!! {label} FAILED (exit {r.returncode})")
            print(r.stderr[-1200:])
            sys.exit(1)

    print("\n" + "=" * 64)
    print("PIPELINE COMPLETE")
    print("=" * 64)
    for label, dt, _ in timings:
        print(f"  {label:<34} {dt:6.2f}s")
    print(f"  {'TOTAL':<34} {time.time()-t0:6.2f}s")
    print("\nDeliverables: ECL_Report.xlsx  +  validation_report.xlsx")

if __name__ == "__main__":
    main()