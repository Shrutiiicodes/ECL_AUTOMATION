"""
main.py - one-command ECL pipeline
==================================
Runs the whole quarterly ECL end to end, in memory, and writes only the two
deliverables:  ECL_Report.xlsx  and  validation_report.xlsx.

    python main.py

Design
------
Every phase is a pure function that takes DataFrames and returns DataFrames; this
orchestrator imports them and chains the results directly - no intermediate CSVs,
no shelling out to scripts. It logs progress per phase, stops on the first
failure, and exits non-zero if validation reports a hard FAIL (so it can gate a
scheduler / CI job).

Phase 0 (synthetic data setup via base_loans.py) is NOT part of the ECL
computation - it only exists because we have no real data. It runs automatically
when ecl.db is missing and is skipped otherwise. For a real cutover, drop it and
point config.DB_PATH at the bank warehouse; nothing downstream changes.
"""

import logging
import os
import subprocess
import sys
import time

from config import DB_PATH, REPORT_XLSX, VALIDATION_XLSX, SEGMENT

import sql_refactor
import chain_ladder
import loss_rate
import final_ecl
import report
import validation

log = logging.getLogger("ecl")


class step:
    """Context manager that logs 'Running <label> ...' and its elapsed time."""
    def __init__(self, label):
        self.label = label

    def __enter__(self):
        self.t0 = time.time()
        log.info("Running %s ...", self.label)
        return self

    def __exit__(self, exc_type, exc, tb):
        dt = time.time() - self.t0
        if exc_type is None:
            log.info("   %s done  (%.2fs)", self.label, dt)
        else:
            log.error("   %s FAILED after %.2fs: %s", self.label, dt, exc)
        return False   # never swallow the exception


def ensure_database():
    """Phase 0: generate the synthetic SQLite DB if it isn't there yet."""
    if os.path.exists(DB_PATH):
        log.info("database %s present - skipping synthetic-data setup", DB_PATH)
        return
    log.info("Running Phase 0  synthetic data setup (base_loans.py) ...")
    r = subprocess.run([sys.executable, "base_loans.py"], capture_output=True, text=True)
    if r.returncode != 0:
        log.error("Phase 0 FAILED:\n%s", r.stderr[-1500:])
        sys.exit(1)
    log.info("   Phase 0 done")


def main():
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s  %(levelname)-5s  %(message)s",
        datefmt="%H:%M:%S",
    )
    t0 = time.time()
    log.info("=" * 56)
    log.info("ECL AUTOMATION - full pipeline")
    log.info("=" * 56)

    ensure_database()

    try:
        with step("Phase 1  SQL refactor"):
            out = sql_refactor.run()
        with step("Phase 2  chain ladder"):
            tris = chain_ladder.run(out.feed, SEGMENT)
        with step("Phase 3  movements + loss rate"):
            lrr = loss_rate.run(tris.a90, tris.atp, out.feed)
        with step("Phase 4  final ECL"):
            ecl = final_ecl.run(lrr.loss, tris.atp)
        with step("Phase 5  Excel report"):
            report.build_excel(out.feed, tris, lrr, ecl, REPORT_XLSX)
        with step("Phase 6  validation"):
            checks = validation.validate(out.feed, tris, lrr, ecl, DB_PATH)
            validation.write_report(checks, VALIDATION_XLSX)
    except Exception as exc:
        log.error("PIPELINE ABORTED: %s", exc)
        sys.exit(1)

    status = validation.overall_status(checks)

    log.info("=" * 56)
    log.info("PIPELINE COMPLETE  (%.2fs)", time.time() - t0)
    log.info("headline window      : %s", ecl.wavg.iloc[0].WINDOW)
    log.info("weighted-avg LR      : %.4f%%", ecl.wavg.iloc[0].WEIGHTED_AVG_LR * 100)
    log.info("ECL (headline)       : %.4f%%", ecl.ecl_pct * 100)
    log.info("deliverables         : %s  +  %s", REPORT_XLSX, VALIDATION_XLSX)
    log.info("validation           : %s", status)
    log.info("=" * 56)

    # full check table
    validation.print_summary(checks)

    if status == "FAIL":
        log.error("validation reported FAIL - do not ship this run")
        sys.exit(1)


if __name__ == "__main__":
    main()