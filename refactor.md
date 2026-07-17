# ECL Automation — Pipeline Refactor

## Summary

The pipeline now runs as a single import-based chain from the SQLite database to the
two Excel deliverables — `python main.py`, no manual SQL step and no intermediate
files. Each of the six phases is a pure, independently testable function; an
independent validation harness recomputes the numbers from the raw tables and
reconciles against the pipeline output.

```
python main.py
  → connects to ecl.db
  → SQL summary (feed) → chain-ladder triangles → movements + loss rates
  → weighted-average loss rate + provisional ECL
  → ECL_Report.xlsx  +  validation_report.xlsx
```

Runtime ≈ 14s on 60k synthetic loans. Only two files are written; every
intermediate is passed in memory as a DataFrame.

## What changed and why

| Area | Before | After |
|------|--------|-------|
| Data flow | each phase read/wrote its own CSVs | phases return DataFrames; chained in memory |
| Execution | scripts ran top-to-bottom on import | each phase is a function (`run()`, `build_excel()`, `validate()`); nothing runs on import |
| Orchestration | `run_all.py` shelled out to each script | `main.py` imports and chains; logging + timings; stops on first failure; non-zero exit on validation FAIL |
| Logic vs reporting | some calc modules also wrote Excel | calc modules return data; `report.py` owns report assembly |
| Real-data cutover | swap the CSV source | pass a different `db_path` to `sql_refactor.run()`; adjust SQL dialect |

Configuration (`AS_OF`, windows, anchors, file names) was already centralised in
`config.py`; that stayed as-is.

## Faithfulness

Every module was verified against the original output before and after the change:
the SQL feed, all four triangles, movement tables, loss rates, and the 8-sheet
workbook are identical (27,200 cells checked, 0 differences). The only values that
moved are two cells at the ~1e-17 level: the in-memory path keeps full float64
precision instead of truncating to 6 decimals on CSV write, which is why the
internal-consistency validation checks now reconcile to exactly `0.00e+00`.

## Standalone use preserved

Each phase can still be run on its own (`python loss_rate.py`, etc.). In standalone
mode it reads the upstream CSVs from disk, runs the same pure function, and writes
its own CSV/Excel artifacts for inspection — useful for debugging a single phase.
Only `main.py` avoids disk entirely.

## Open modelling question (needs mentor input)

The 84M loss rate reconciles to the hand calculation exactly. The 120M window
reports 146.7%, which is impossible: only 5 of 32 window cohorts are mature at
120M, and the Excel chain-ladder formula compounds geometrically down each column,
so the 27 immature cohorts inflate. Restricted to genuinely-mature cohorts the
120M rate is 2.41%, consistent with 84M's mature-only 2.50%.

`validation.py` flags any reported loss rate above 100% as **WARN** rather than
hiding it. Before the 120M figure can be used, the production behaviour needs
confirming: does the workbook project deep columns at all, cap them, or restrict
the 120M window to mature cohorts only?

## File map

```
config.py            single source of truth (AS_OF, windows, anchors, filenames)
base_loans.py        Phase 0 — synthetic data → ecl.db  (dropped for real data)
sql_refactor.py      run() → feed DataFrame from SQL
chain_ladder.py      run(feed) → Triangles(r90, a90, rtp, atp, mat90, mattp, disb, feed)
loss_rate.py         run(a90, atp, feed) → LossRates(loss, mv90, mvtp)
final_ecl.py         run(loss, atp) → FinalECL(by_quarter, wavg, portfolio_tpos)
report.py            build_excel(feed, tris, lrr, ecl) → ECL_Report.xlsx
validation.py        validate(feed, tris, lrr, ecl) → [Check]; overall_status(); write_report()
main.py              orchestrator: chains the above in memory, logging + stop-on-failure
```