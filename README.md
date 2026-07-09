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

Replaces the manual quarterly process (run SQL, paste to Excel, build pivots, mark yellow cells, drag chain-ladder formulas, build movement tables, compute loss rates, take the weighted average) with a single run that produces a formatted `ECL_Report.xlsx` and a `validation_report.xlsx` proving the run reconciles.

## One-command run

```
python run_all.py
```

Runs the full pipeline in order, stops on the first failure, prints timings. ≈ 30s on 60k loans.

**Deliverables:** `ECL_Report.xlsx` (8 tabs) and `validation_report.xlsx` (must be all-PASS before you ship).

## Phase map

| Phase | Script | Produces |
|------|--------|----------|
| 0 | `base_loans.py` | `ecl.db` — raw `base_loans` + long `performance` tables |
| 1 | `sql_refactor.py` | `DATA_ECL_NEW.csv` — generated SQL replaces ~82 joins; `phase1_generated.sql` |
| 3 | `chain_ladder.py` | completed 90+/TPOS triangles (Phase 2 folded in) |
| 4 | `loss_rate.py` | movement tables + per-quarter loss rate at **84M and 120M** |
| 5 | `final_ecl.py` | **disbursal-weighted average loss rate** per observation window |
| 6 | `report.py` | `ECL_Report.xlsx` |
| 7 | `validation.py` | `validation_report.xlsx` (independent reconciliation) |

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

reproducing `=SUMPRODUCT(O125:O140,$B125:$B140)/SUM($B125:$B140)`. Weights are disbursal amounts. Only the FY range and anchor change per window. Configured in `final_ecl.py`:

| Window | Range | Anchor |
|---|---|---|
| FY20–FY23 @ 84M | FY20-Q1 → FY23-Q4 | 84M |
| FY16–FY23 @ 84M | FY16-Q1 → FY23-Q4 | 84M |
| FY16–FY23 @ 120M | FY16-Q1 → FY23-Q4 | 120M |

**Provisional final step:** `ECL = headline weighted_LR × portfolio current TPOS`, where current TPOS is the latest observed outstanding (triangle diagonal). The weighted rate is per spec; the multiplier is not yet confirmed.

## Changing the quarter

Edit `config.py`: set `AS_OF` and the `START_DISB`/`END_DISB` window. `AS_OF` drives cohort maturity (which cells are projected) and the current-TPOS diagonal. Windows are set in `final_ecl.py` (`WINDOWS`, `HEADLINE`).

> The phase scripts carry the same values in an inline `CONFIG` block for standalone clarity, so a run works out of the box. Recommended one-step refactor: replace each inline block with `from config import *`.

## Requirements

Python 3.10+ with `pandas`, `numpy`, `openpyxl`, `python-dateutil`. `sqlite3` is built in. No database server needed — the pipeline uses a local `ecl.db` file.

## Real-data cutover

Today `base_loans.py` generates synthetic data. To run on real bank data, drop that stage and point `sql_refactor.py` at the actual `base_loans` + `performance` tables (Netezza/Teradata). The generated SQL ports with dialect tweaks: `strftime` → the platform's date functions, `1e7` division unchanged. Everything downstream is unchanged.

## Known issue: the 120M anchor is unreliable

The Excel chain-ladder formula compounds **down** each MOB column — every projected cell is the row above scaled by a factor > 1, so each projection feeds the next.

At **84M** this is fine: 17 of the 32 window cohorts are fully mature, projections are short, and the weighted rate lands at a plausible **3.75%** (FY16–FY23) / **5.03%** (FY20–FY23).

At **120M** only **5 of 32** cohorts are mature. The remaining 27 rows compound geometrically — the per-cohort rate climbs 2.4% → 2.95% → 3.71% → 4.84% → 6.13% → 9.01% → … and the weighted average reaches **146.7%**, which is impossible.

Restricted to cohorts genuinely mature at 120M, the rate is **2.41%** — consistent with 84M's mature-only 2.50%.

`validation.py` flags any reported loss rate above 100% as **WARN**. Before using the 120M figure, confirm with the mentor: does the production workbook project deep columns at all, cap them, or restrict the 120M window to mature cohorts only?

## Validation philosophy

Never trust automation blind. `validation.py` recomputes every stage independently (not reusing pipeline code) and diffs against the outputs with a crore-level tolerance, plus a plausibility check. Ship only when all checks read PASS.