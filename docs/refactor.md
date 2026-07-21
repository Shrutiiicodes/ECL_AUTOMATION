# ECL Automation — Pipeline Refactor

## Summary

The pipeline runs as a single import-based chain from the SQLite database to the two Excel deliverables — `python main.py`, with no manual SQL step and no intermediate files. Each computation phase is a pure, independently testable function, and two independent validation layers reconcile the results before the run is trusted.

```
python main.py
  → (Phase 0) generate data/db/ecl.db if missing
  → (Phase 1) SQL summary feed
  → (Phase 2) chain-ladder triangles (90+ and TPOS)
  → (Phase 3) movement tables + per-anchor loss rates
  → (Phase 4) disbursal-weighted average loss rate + headline ECL %
  → (Phase 5) ECL_Report.xlsx
  → (Phase 6) validation_report.xlsx  — independent numeric reconciliation
  → (Phase 7) excel-layer validation  — structural + optional recalc of the workbook
```

Only two files are written (`output/ECL_Report.xlsx` and `output/validation_report.xlsx`); every intermediate is passed in memory as a DataFrame. Runtime is ≈ 15s on 60k synthetic loans, or ≈ 30s when Phase 7 has LibreOffice available for the Level-2 recalc.

## What changed and why

| Area | Before | After |
|------|--------|-------|
| Data flow | each phase read/wrote its own CSVs | phases return DataFrames; chained in memory |
| Execution | scripts ran top-to-bottom on import | each phase is a function (`run()`, `build_excel()`, `validate()`, `run_excel_validation()`); nothing runs on import |
| Orchestration | `run_all.py` shelled out to each script | `main.py` imports and chains; logging + timings; stops on first failure; non-zero exit on validation FAIL |
| Logic vs reporting | some calc modules also wrote Excel | calc modules return data; `report.py` owns report assembly |
| Layout | flat scripts in the repo root | `src/` (phases + config), `scripts/` (synthetic data), `docs/` |
| Validation | numeric reconciliation only | numeric reconciliation (Phase 6) **plus** a check of the delivered workbook itself (Phase 7) |
| Real-data cutover | swap the CSV source | pass a different `db_path` to `sql_refactor.run()`; adjust SQL dialect |

Configuration (`AS_OF`, windows, anchors, file names) is centralised in `src/config.py`. Every phase module imports it with `from src.config import *` and defines no shared constants of its own.

## Phase functions

```
src/config.py            single source of truth (AS_OF, windows, anchors, filenames)
scripts/base_loans.py    Phase 0 — synthetic data → data/db/ecl.db  (dropped for real data)
src/sql_refactor.py      run() → SqlOutput(feed, sql)
src/chain_ladder.py      run(feed, segment) → Triangles(r90, a90, rtp, atp, mat90, mattp, disb, feed)
src/loss_rate.py         run(a90, atp, feed) → LossRates(loss, mv90, mvtp)
src/final_ecl.py         run(loss, atp) → FinalECL(by_quarter, wavg, portfolio_tpos, ecl_pct)
src/report.py            build_excel(feed, tris, lrr, ecl, path) → ECL_Report.xlsx (7 tabs)
src/validation.py        validate(feed, tris, lrr, ecl, db_path) → [Check]; overall_status(); write_report()
src/excel_validation.py  run_excel_validation(ecl_pct, path) → "PASS"/"FAIL"  (Level 1 always, Level 2 if LibreOffice)
main.py                  orchestrator: chains the above in memory, logging + stop-on-failure
```

## Standalone use preserved

Each phase can still be run on its own (`python -m src.loss_rate`, etc.). In standalone mode it reads the upstream CSVs from disk, runs the same pure function, and writes its own CSV/Excel artifacts (and prints a self-check) for inspection — useful for debugging a single phase. Only `main.py` avoids disk entirely. Because `main.py` keeps everything in memory, the intermediate CSVs under `data/intermediate/` only appear when a phase is run standalone.

## Loss-rate denominator

The loss rate at anchor A now divides 90+ by the **disbursal amount plus** the TPOS at the yearly MOB levels up to one year less than the anchor:

```
LR_A(q) = 90+(q, A) / ( DISBURSAL_AMT(q) + SUM( TPOS(q,12), ..., TPOS(q, A-12) ) )
```

In the workbook this is a single contiguous `SUM(Movements!B:<A-12>)` (B = disbursal, C onward = TPOS levels). Anchors are configured in `ANCHOR_MOBS` (currently 72M, 84M, 120M); `validation.py` runs one reconciliation check per anchor.

## Chain-ladder fill

Projected cells use the bank's Excel development-factor formula, `=IFERROR(F46*SUMPRODUCT(G$4:G45,$B$4:$B45)/SUMPRODUCT(F$4:F45,$B$4:$B45),0)`: the base is the same cohort's previous-MOB cell and the multiplier is the disbursal-weighted development ratio from the previous MOB to this MOB, over the mature cohorts above. The 90+ triangle is filled in rate space then converted to amount; the TPOS triangle is filled in amount space. The numpy engine (`chain_ladder.chain_ladder_fill`) and the live formulas emitted by `report.py` match cell-for-cell.

## The 120M anchor

With the disbursal amount in the denominator, all three anchors currently produce plausible loss rates (a low single-digit percentage on the synthetic data, none exceeding 100%). Earlier iterations — before disbursal was added to the denominator — could report >100% at 120M, where far fewer cohorts are mature. That is no longer observed. `validation.py` keeps a plausibility check that flags any reported weighted-average loss rate above 100% as **WARN** (not FAIL), as a safety net for real data.

## Faithfulness

Every stage is reconciled on every run, so a broken change fails loudly rather than shipping silently:

- **Numeric layer (Phase 6):** the SQL feed, triangles, movement tables, per-anchor loss rates, and weighted averages are recomputed independently from the raw DB tables and diffed against the pipeline output to `TOL = 1e-4`. Because the in-memory path keeps full float64 precision (instead of truncating to 6 decimals on CSV write), the internal-consistency checks reconcile to `0.00e+00`.
- **Workbook layer (Phase 7):** the delivered `ECL_Report.xlsx` is checked structurally (live formulas wired to the right cells), and — when LibreOffice is available — recalculated headless and reconciled back against the Python engine to `TOL`.

A run is safe to ship only when both layers read PASS.