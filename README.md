# ECL Automation

## What ECL is

ECL stands for **Expected Credit Loss** — it sits at the intersection of accounting, credit risk, and data. It is the amount a bank expects to lose on its loans because some borrowers won't fully pay back. Instead of waiting for a loan to go bad and only then booking a loss, the bank estimates losses in advance and sets aside provisions for them. Every loan carries some probability of not being repaid; ECL puts a forward-looking number on that risk today, rather than reacting after the damage is done. This framework replaced the earlier "incurred loss" model after the 2008 financial crisis.

The current accounting standard is IFRS 9 in Europe, with CECL as the US equivalent. Both are principles-based rather than rules-based, so they leave room for judgement.

Conceptually the calculation is simple:

```
ECL = Probability of Default (PoD) × Loss Given Default (LGD) × Exposure at Default (EAD)
```

where PoD is the likelihood the borrower defaults, LGD is how much the bank loses if they do (net of recovery), and EAD is how much is owed at the point of default. EAD is usually just the outstanding balance; the complexity lives in estimating PoD and LGD. This project estimates loss empirically from historical run-off (a chain-ladder loss-rate approach) rather than modelling PoD/LGD separately.

---

## Run-book

End-to-end automation of the quarterly PL (NON-PA) ECL: **SQL → Python → Excel**, in one command.

### What it does

Replaces the manual quarterly process — run SQL, paste into Excel, build pivots, mark the immature (yellow) cells, drag chain-ladder formulas, build movement tables, compute loss rates, take the weighted average — with a single run that produces a formatted `ECL_Report.xlsx` (7 tabs) plus a `validation_report.xlsx` proving the run reconciles.

### One-command run

```
python main.py
```

The whole pipeline runs **in memory**: every phase is a pure function that takes DataFrames and returns DataFrames, and the orchestrator chains the results directly — no intermediate CSVs, no shelling out to per-phase scripts. It logs progress and timings per phase, stops on the first failure, and exits non-zero if validation reports a hard FAIL (so it can gate a scheduler or CI job). Runtime is ≈ 15s on 60k synthetic loans, or ≈ 30s when the optional Excel-recalc validation (Phase 7, Level 2) has LibreOffice available.

**Deliverables:** `output/ECL_Report.xlsx` (7 tabs) and `output/validation_report.xlsx` (should read all-PASS before you ship).

### Architecture

All configuration lives in a single file, `src/config.py`. Every phase module begins with `from src.config import *` and defines no shared constants of its own — to roll the quarter you change `AS_OF` and nothing else.

Each phase module exposes a **pure function** that takes DataFrames in and returns a `NamedTuple` out, with no file I/O. The `if __name__ == "__main__"` block in each module exists only so that phase can still be run and inspected standalone (reading the upstream CSVs from disk and writing its own artifacts); none of that runs on import.

`main.py` imports the phase modules and chains their results:

```python
ensure_database()                                  # Phase 0 — only if data/db/ecl.db is missing
out    = sql_refactor.run()                         # Phase 1
tris   = chain_ladder.run(out.feed, SEGMENT)        # Phase 2
lrr    = loss_rate.run(tris.a90, tris.atp, out.feed) # Phase 3
ecl    = final_ecl.run(lrr.loss, tris.atp)          # Phase 4
report.build_excel(out.feed, tris, lrr, ecl, REPORT_XLSX)          # Phase 5
checks = validation.validate(out.feed, tris, lrr, ecl, DB_PATH)   # Phase 6
validation.write_report(checks, VALIDATION_XLSX)
xl_status = excel_validation.run_excel_validation(ecl.ecl_pct, REPORT_XLSX)  # Phase 7
```

### Phase map

| Phase | Module | Pure function | Produces |
|-------|--------|---------------|----------|
| 0 | `scripts/base_loans.py` | *(standalone script, run via `python -m scripts.base_loans`)* | `data/db/ecl.db` — synthetic `base_loans` + long `performance` tables |
| 1 | `src/sql_refactor.py` | `run() → SqlOutput(feed, sql)` | `data_ecl` feed (one row per FY_QUARTER); generated SQL replaces ~82 joins |
| 2 | `src/chain_ladder.py` | `run(feed, segment) → Triangles(r90, a90, rtp, atp, mat90, mattp, disb, feed)` | Completed 90+ / TPOS triangles (rate and amount) |
| 3 | `src/loss_rate.py` | `run(a90, atp, feed) → LossRates(loss, mv90, mvtp)` | Movement tables + per-quarter loss rate at each anchor (72M, 84M, 120M) |
| 4 | `src/final_ecl.py` | `run(loss, atp) → FinalECL(by_quarter, wavg, portfolio_tpos, ecl_pct)` | Disbursal-weighted average loss rate per observation window |
| 5 | `src/report.py` | `build_excel(feed, tris, lrr, ecl, path) → Workbook` | `ECL_Report.xlsx` (7 tabs) |
| 6 | `src/validation.py` | `validate(feed, tris, lrr, ecl, db_path) → list[Check]` | `validation_report.xlsx` (independent numeric reconciliation) |
| 7 | `src/excel_validation.py` | `run_excel_validation(ecl_pct, path) → "PASS"/"FAIL"` | Console report validating the delivered **workbook** (structure + optional recalc) |

Phase 0 is **not** part of the ECL computation — it generates synthetic data because there is no real bank data available. It runs automatically only when `data/db/ecl.db` is missing, and is skipped otherwise.

### ECL_Report.xlsx tabs

Every computed cell in the workbook is a **live Excel formula** (`fullCalcOnLoad` is set so results show the moment it opens). Disbursal amounts are a single source of truth: every `DISB` cell is a `SUMIF` back to the `DATA_ECL` sheet keyed on `FY_QUARTER`.

| # | Tab | Contents |
|---|-----|----------|
| 1 | Summary | As-of date, cohort count, MOB grid, loss-rate anchors, headline window, headline weighted-avg loss rate, window total disbursal, ECL %, and the final-ECL rule |
| 2 | DATA_ECL | The raw SQL feed, one row per FY_QUARTER (values) |
| 3 | Pivot_ECL | RAW pivot: 90+ amount block + TPOS amount block, observed actuals only (immature cells blank, no yellow), each with a Grand Total |
| 4 | Chain_Ladder | Chain-ladder triangles with live formulas: 90+ as % and TPOS as amount. Mature cells link to Pivot_ECL; immature (yellow) cells carry the development-factor formula |
| 5 | Movements | TPOS movement (amount), TPOS movement (% of disbursal), 90+ movement (% of disbursal) at the yearly MOB levels |
| 6 | LossRate | Per-quarter loss rate at 72M / 84M / 120M, plus CURRENT_MOB and CURRENT_TPOS. Projected anchors flagged yellow; observed-but-implausible (>100%) flagged red |
| 7 | Weighted_LR | `SUMPRODUCT(LR, DISB)/SUM(DISB)` per observation window, plus the headline ECL % and its window disbursal |

### The calculation

**Per-quarter loss rate at anchor A**

```
LR_A(q) = 90+(q, A) / ( DISBURSAL_AMT(q) + SUM( TPOS(q,12), TPOS(q,24), ..., TPOS(q, A-12) ) )
```

The numerator is the 90+ settlement amount at MOB A. The denominator is the cohort's **disbursal amount plus** the TPOS at every yearly MOB up to **one year less than the anchor** (72M → up to 60, 84M → up to 72, 120M → up to 108). In the workbook this denominator is read as one contiguous range off the Movements sheet — `SUM(Movements!B:<A-12>)`, where column B is disbursal and columns C onward are the TPOS levels — so 72M is `SUM(B:G)`, 84M is `SUM(B:H)`, and 120M is `SUM(B:K)`. Anchors are MOB **levels** at those ages, not deltas. Supported anchors are configured in `ANCHOR_MOBS` and are currently 72M, 84M, and 120M.

**Weighted-average loss rate over an observation window**

```
weighted_LR = SUMPRODUCT(LR[q1:q2], DISB[q1:q2]) / SUM(DISB[q1:q2])
```

reproducing `=SUMPRODUCT(O125:O140,$B125:$B140)/SUM($B125:$B140)`. The weights are disbursal amounts; only the FY range and the anchor change per window. Windows are configured in `src/config.py` (`WINDOWS`, `HEADLINE`):

| Window | Range | Anchor |
|---|---|---|
| FY20-FY23 @ 84M *(headline)* | FY20-Q1 → FY23-Q4 | 84M |
| FY16-FY23 @ 84M | FY16-Q1 → FY23-Q4 | 84M |
| FY16-FY23 @ 120M | FY16-Q1 → FY23-Q4 | 120M |

**Final step.** The reported **ECL % is the headline window's weighted-average loss rate itself**. `run()` returns that rate as `ecl_pct`; the headline window's total disbursal is carried alongside (and shown on the Summary sheet) as the exposure the rate would apply to, but the pipeline currently reports the rate, not a rupee amount. `CURRENT_TPOS` (the latest observed outstanding, i.e. the triangle diagonal) is computed and shown for information only — it is not used as the exposure.

### Chain-ladder fill

For an immature (projected) cell at row `R`, MOB column `X`, with previous MOB column `P`:

```
val(R, X) = val(R, P) × SUMPRODUCT(X[top..R-1], DISB[top..R-1]) / SUMPRODUCT(P[top..R-1], DISB[top..R-1])
```

The base is the **same cohort's previous-MOB value** and the multiplier is the disbursal-weighted development ratio from the previous MOB to this MOB, taken over the mature cohorts **above** this row. This reproduces the manual workbook's `=IFERROR(F46*SUMPRODUCT(G$4:G45,$B$4:$B45)/SUMPRODUCT(F$4:F45,$B$4:$B45),0)` cell-for-cell — both in the numpy engine (`chain_ladder.chain_ladder_fill`) and in the live formulas emitted by `report.py`. The 90+ triangle is filled in **rate** space (% cells) and then converted to amount = rate × disbursal; the TPOS triangle is filled directly in **amount** space. `IFERROR → 0` where the denominator is 0 or there is no previous column / no row above.

### Changing the quarter

Edit `src/config.py` and set `AS_OF` (the single knob). `END_DISB` is derived from `AS_OF` so the two can never silently disagree. `AS_OF` drives cohort maturity (which cells are projected) and the current-TPOS diagonal. Observation windows and the headline window are also in `src/config.py` (`WINDOWS`, `HEADLINE`).

### MOB grids

Two grids are kept deliberately separate:

- `MOB_SQL` = 0, 3, …, 120 (41 points) — what the extraction layer **captures** (0MOB is captured because the bank's extract captures it).
- `MOB_LIST` = 3, 6, …, 120 (40 points) — the **pivot / triangle** grid, per spec. 0MOB is never a pivot column, so it is never chain-laddered.

### File structure

```
ECL_AUTOMATION/
├── main.py                  ← entry-point orchestrator (Phases 0–7, in memory)
├── src/
│   ├── config.py            ← single source of truth for all shared knobs
│   ├── sql_refactor.py      ← Phase 1: generated SQL replacing ~82 joins
│   ├── chain_ladder.py      ← Phase 2: triangle build + chain-ladder fill
│   ├── loss_rate.py         ← Phase 3: movement tables + per-anchor loss rates
│   ├── final_ecl.py         ← Phase 4: weighted-average loss rate
│   ├── report.py            ← Phase 5: ECL_Report.xlsx (7 tabs, live formulas)
│   ├── validation.py        ← Phase 6: independent numeric reconciliation
│   ├── excel_validation.py  ← Phase 7: structural + recalc check of the workbook
│   └── __init__.py
├── scripts/
│   └── base_loans.py        ← Phase 0: synthetic data generation → ecl.db
├── docs/
│   └── refactor.md          ← notes on the in-memory pipeline refactor
├── data/                    ← git-ignored: db/ecl.db and intermediate CSVs
├── output/                  ← git-ignored: ECL_Report.xlsx, validation_report.xlsx
├── requirements.txt         ← numpy, pandas, python-dateutil, openpyxl
└── .gitignore               ← ignores data/, output/, reference docs/, *.csv, *.xlsx, ecl.db
```

Everything under `data/` and `output/` is regenerated by `python main.py` and git-ignored.

### Requirements

Python 3.10+ with `pandas`, `numpy`, `openpyxl`, and `python-dateutil`. `sqlite3` is built in. No database server is needed — the pipeline uses a local `data/db/ecl.db` file.

```
pip install -r requirements.txt
```

LibreOffice is **optional**: Phase 7 uses it to recalculate the delivered workbook headless and reconcile the computed values against the Python engine (Level 2). If LibreOffice is not installed, Phase 7 still runs its structural checks (Level 1) and reports SKIP for the recalc — it never fails just because LibreOffice is absent.

### Real-data cutover

Today `scripts/base_loans.py` generates synthetic data. To run on real bank data, drop that stage and point `config.DB_PATH` (or the `db_path` argument of `sql_refactor.run()`) at the actual `base_loans` + `performance` tables (e.g. Netezza / Teradata). The generated SQL ports with dialect tweaks: `strftime` → the platform's date functions, the `/1e7` crore division unchanged. Everything downstream is unchanged.

### Loss rates and the 120M anchor

Because the loss-rate **denominator includes the disbursal amount** (`DISB + SUM(TPOS 12..A-12)`), reported rates are a small fraction of disbursal, and all three anchors currently produce plausible, similar figures. On the bundled synthetic data the headline (FY20-FY23 @ 84M) weighted-average loss rate is ≈ 1.2%, and the 84M and 120M windows over FY16-FY23 land at essentially the same value — none exceed 100%. (Earlier iterations, before the disbursal amount was added to the denominator, could produce implausible >100% figures at 120M, where far fewer cohorts are mature; that is no longer observed.)

`validation.py` retains a plausibility check as a safety net: any reported weighted-average loss rate above 100% is flagged **WARN** (not FAIL). If real data ever trips it, confirm with the mentor how the production workbook treats deep projected columns before shipping.

### Validation philosophy

Never trust automation blind. There are two independent layers.

`validation.py` (Phase 6) recomputes every stage **independently** — from the raw DB tables, not by reusing pipeline code — and diffs against the pipeline outputs with a crore-level tolerance (`TOL = 1e-4`), plus a plausibility check. There is one loss-rate check per anchor, so the count scales with `ANCHOR_MOBS`; with `[72, 84, 120]` there are nine checks:

| # | Check | Method |
|---|-------|--------|
| 1 | LAN_CNT | vs pandas groupby on raw DB tables (exact) |
| 2 | DISBURSAL_AMT | vs pandas groupby |
| 3 | 90+ sums (all MOB) | vs pandas groupby |
| 4 | TPOS sums (all MOB) | vs pandas groupby |
| 5 | loss_rate 72M | vs recompute from the amount triangles |
| 6 | loss_rate 84M | vs recompute from the amount triangles |
| 7 | loss_rate 120M | vs recompute from the amount triangles |
| 8 | weighted-avg LR | vs explicit SUMPRODUCT/SUM per window |
| 9 | plausibility | every reported weighted LR in (0,1) — WARN, not FAIL |

`excel_validation.py` (Phase 7) validates the **delivered workbook** itself, which the numeric harness never opens. Level 1 (always runs) asserts structure: the sheets exist, the headline ECL cell is a live `SUMPRODUCT/SUM` formula rather than a pasted literal, and the Summary ECL % cell links into the weighted-LR column of `Weighted_LR` (catching wiring bugs where a correct calculation points at the wrong cell). Level 2 (only if LibreOffice is present) recalculates the workbook headless and reconciles the computed ECL back against the Python engine to `TOL`. Ship only when both layers read PASS.