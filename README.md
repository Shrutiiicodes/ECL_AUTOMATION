ECL AUTOMATION:
ECL stands for Expected Credit Loss. It sits at the intersection of accounting, credit risk and data.
It is essentially the amount of money that a bank expects to lose on its loans as some borrowers wont fully pay back.
So instead of waiting for a loan to go bad and then book a loss, the bank can estimate the losses in advance and set aside provision for them.
The core idea is that every loan carries a certain probability of not being repaid. ECL puts a number on that risk today forward-looking rather than reacting after the damage is done already. This framework replaced the earlier model which was "incurred loss" after the 2008 financial crisis.

Currently, the accounting standard is IFRS 9 in Europe, while the US equivalent is CECL. Both are principles based rather than rules based so they still leave room for judgement.

The ECL calculation itself is actually quite simple conceptually:
ECL = Probability of Default × Loss Given Default × Exposure at Default

POD: likelihood the borrower will default
LGD: how much the bank expects to lose if they default (recovery rate)
EAD: how much the borrower owes at the point of default (loan balance)

Where the complexity comes in is calculating these three numbers.
EAD is usually straightforward - the outstanding loan balance. The hard part is POD and LGD.
# ECL Automation — Run-book

End-to-end automation of the quarterly PL (NON-PA) ECL: **SQL → Python → Excel**, one command.

## What it does

Replaces the manual quarterly process (run SQL, paste to Excel, build pivots, mark yellow cells, drag chain-ladder formulas, build movement tables, compute loss rates, take the weighted average) with a single run that produces a formatted `ECL_Report.xlsx` (8 tabs) and a `validation_report.xlsx` proving the run reconciles.

## One-command run

```
python main.py
```

Runs the full pipeline **in memory** — every phase is a pure function that takes DataFrames and returns DataFrames; the orchestrator chains the results directly with no intermediate CSVs and no shelling out to scripts. Stops on the first failure, prints timings and a validation summary. ≈ 30s on 60k loans.

**Deliverables:** `ECL_Report.xlsx` (8 tabs) and `validation_report.xlsx` (must be all-PASS before you ship).

## Architecture

All configuration lives in a single file: `config.py`. Every phase script begins with `from config import *` and defines no shared constants of its own. To roll the quarter you change `AS_OF` in `config.py` and nothing else.

Each phase module exposes a **pure function** (`run()` or equivalent) that takes DataFrames in and returns NamedTuples out — no file I/O. The `__main__` block in each script exists only so that phase can still be run and inspected standalone; none of it executes on import.

`main.py` imports all phase modules and chains the results:

```
ensure_database()          # Phase 0 — only if ecl.db is missing
out  = sql_refactor.run()
tris = chain_ladder.run(out.feed, SEGMENT)
lrr  = loss_rate.run(tris.a90, tris.atp, out.feed)
ecl  = final_ecl.run(lrr.loss, tris.atp)
report.build_excel(out.feed, tris, lrr, ecl, REPORT_XLSX)
validation.validate(out.feed, tris, lrr, ecl, DB_PATH)
```

## Phase map

| Phase | Script | Pure function | Produces |
|-------|--------|---------------|----------|
| 0 | `base_loans.py` | *(standalone script)* | `ecl.db` — synthetic `base_loans` + long `performance` tables |
| 1 | `sql_refactor.py` | `run() → SqlOutput(feed, sql)` | `data_ecl` summary (one row per FY_QUARTER × SEGMENT); generated SQL replaces ~82 joins |
| 2 | `chain_ladder.py` | `run() → Triangles(r90, a90, rtp, atp, mat90, mattp, disb, feed)` | Completed 90+ / TPOS triangles |
| 3 | `loss_rate.py` | `run() → LossRates(loss, mv90, mvtp)` | Movement tables + per-quarter loss rate at **84M** and **120M** |
| 4 | `final_ecl.py` | `run() → FinalECL(by_quarter, wavg, portfolio_tpos)` | Disbursal-weighted average loss rate per observation window |
| 5 | `report.py` | `build_excel()` | `ECL_Report.xlsx` |
| 6 | `validation.py` | `validate() → list[Check]` | `validation_report.xlsx` (independent reconciliation) |

Phase 0 is **not** part of the ECL computation — it generates synthetic data because we have no real bank data. It runs automatically when `ecl.db` is missing and is skipped otherwise.

## ECL_Report.xlsx tabs

| # | Tab | Contents |
|---|-----|----------|
| 1 | Summary | Cover metrics + headline weighted loss rate |
| 2 | data_ecl | The SQL feed (FY_QUARTER × SEGMENT) |
| 3 | Pivot_90plus | 90+ amount triangle, yellow = chain-ladder projected |
| 4 | Pivot_TPOS | TPOS amount triangle, yellow = projected |
| 5 | BadRate_90plus | 90+ / DISB rate triangle (PD curve), yellow = projected |
| 6 | Movements | TPOS + 90+ movement tables (12…120) |
| 7 | LossRate_Qtr | Per-quarter loss rate at 84M and 120M (>100% flagged red) |
| 8 | Weighted_LR | SUMPRODUCT(LR, DISB)/SUM(DISB) per observation window |

## The calculation

**Per-quarter loss rate at anchor A**

```
LR_A(q) = 90+(q, A) / SUM( TPOS(q,12), TPOS(q,24), ..., TPOS(q,A) )
```

reproducing `=I97/SUM(B51:H51)` (I = 90+ at 84M, B:H = TPOS movement 12→84). Anchors are MOB **levels**, not deltas. Supported anchors: 84M (7 terms) and 120M (10 terms).

**Weighted-average loss rate over an observation window**

```
weighted_LR = SUMPRODUCT(LR[q1:q2], DISB[q1:q2]) / SUM(DISB[q1:q2])
```

reproducing `=SUMPRODUCT(O125:O140,$B125:$B140)/SUM($B125:$B140)`. Weights are disbursal amounts. Only the FY range and anchor change per window. Configured in `config.py` (`WINDOWS`, `HEADLINE`):

| Window | Range | Anchor |
|---|---|---|
| FY16–FY23 @ 84M | FY16-Q1 → FY23-Q4 | 84M |
| FY16–FY23 @ 120M | FY16-Q1 → FY23-Q4 | 120M |

**Final step:** The weighted rate is per spec;

## Changing the quarter

Edit `config.py`: set `AS_OF` (the single knob). `END_DISB` is derived from `AS_OF` so the two can never silently disagree. `AS_OF` drives cohort maturity (which cells are projected) and the current-TPOS diagonal. Observation windows are also set in `config.py` (`WINDOWS`, `HEADLINE`).

## File structure

```
ECL-AUTOMATION/
├── main.py              ← one-command orchestrator
├── config.py            ← single source of truth for all shared knobs
├── base_loans.py        ← Phase 0: synthetic data generation
├── sql_refactor.py      ← Phase 1: generated SQL replacing ~82 joins
├── chain_ladder.py      ← Phase 2: triangle build + chain-ladder fill
├── loss_rate.py         ← Phase 3: movement tables + loss rates
├── final_ecl.py         ← Phase 4: weighted-average loss rate
├── report.py            ← Phase 5: ECL_Report.xlsx (8 tabs)
├── validation.py        ← Phase 6: independent reconciliation
├── requirements.txt     ← numpy, pandas, python-dateutil, openpyxl
└── .gitignore           ← all .csv/.xlsx/ecl.db are regenerated artifacts
```

Everything under `*.csv`, `*.xlsx`, and `ecl.db` is regenerated by `python main.py` and git-ignored. The `reference from the meeting/` directory (screenshots + legacy scripts) is also git-ignored.

## Requirements

Python 3.10+ with `pandas`, `numpy`, `openpyxl`, `python-dateutil`. `sqlite3` is built in. No database server needed — the pipeline uses a local `ecl.db` file.

```
pip install -r requirements.txt
```

## Real-data cutover

Today `base_loans.py` generates synthetic data. To run on real bank data, drop that stage and point `config.DB_PATH` at the actual `base_loans` + `performance` tables (Netezza/Teradata). The generated SQL ports with dialect tweaks: `strftime` → the platform's date functions, `1e7` division unchanged. Everything downstream is unchanged.

## Note: 120M anchor behaviour

The chain-ladder fill compounds **down** each MOB column — every projected cell is the row above scaled by a disbursal-weighted trend factor. Because the 120M anchor has far fewer mature cohorts than the 84M anchor, its projections span more rows and are inherently noisier. Earlier versions of the pipeline produced implausible 120M rates (>100%) due to unchecked geometric compounding; this has since been fixed and both anchors now produce plausible weighted-average loss rates.

`validation.py` still flags any reported loss rate above 100% as **WARN** as a safety net.

## Validation philosophy

Never trust automation blind. `validation.py` recomputes every stage independently (not reusing pipeline code) and diffs against the outputs with a crore-level tolerance, plus a plausibility check. Ship only when all checks read PASS.

| # | Check | Method |
|---|-------|--------|
| 1 | LAN_CNT | vs pandas groupby on raw DB tables (exact) |
| 2 | DISBURSAL_AMT | vs pandas groupby |
| 3 | 90+ sums (all MOB) | vs pandas groupby |
| 4 | TPOS sums (all MOB) | vs pandas groupby |
| 5 | loss_rate 84M | vs recompute from triangles |
| 6 | loss_rate 120M | vs recompute from triangles |
| 7 | weighted-avg LR | vs explicit SUMPRODUCT/SUM per window |
| 8 | plausibility | every reported weighted LR in (0,1) — WARN, not FAIL |