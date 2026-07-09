"""
PHASE 0 - DATA FOUNDATION
=========================
Generates synthetic ECL data in the BANK'S RAW SHAPE (two tables) and loads it
into a local SQLite database. This is the input that Phase 1 SQL will reassemble.

Tables produced
---------------
base_loans   : one row per distinct loan (loan grain)
                 distinct_loan_no, customer_id, disbursal_date,
                 disbursal_amount (RUPEES), segment, tenure_months
performance  : LONG format, one row per loan x OBSERVED MOB
                 distinct_loan_no, mob, tpos (RUPEES), amt_90plus_settlement (RUPEES)

Key design notes
----------------
* Amounts are stored in RUPEES. Phase 1 SQL applies /10^7 to report in crores,
  exactly like the bank's sum(...)/10^7. Keep the conversion in SQL, not here.
* performance holds ONLY observed rows (mob <= loan age at AS_OF). A loan
  disbursed 7 months ago has rows for mob 0,3,6 only. This asymmetry is what
  produces the run-off triangle (NULLs for immature cells) after the join.
* fy_quarter is NOT stored. Phase 1 SQL derives it from disbursal_date so the
  FY-quarter logic lives in SQL (mirrors the bank's concat/case expression).
* Same seed / distributions / default model as your data_generation.py, so the
  Phase-1 widened output reconciles with your existing wide CSV (x 1e7).

Outputs: ecl.db (SQLite), base_loans.csv, performance_long.csv
"""

import sqlite3
from datetime import date

import numpy as np
import pandas as pd
from dateutil.relativedelta import relativedelta

# CONFIG
SEED       = 34
N_LOANS    = 60_000
START      = date(2015, 4, 1)
AS_OF      = date(2026, 7, 7)          # reporting "today"; drives triangle maturity
MOB_LIST   = list(range(0, 121, 3))    # 0,3,6,...,120  (41 points)
SEGMENTS   = [1, 2, 3, 4, 5]
SEG_DEFAULT_RATE = {1: 0.020, 2: 0.035, 3: 0.055, 4: 0.080, 5: 0.120}

# If True, a loan's TPOS freezes at outstanding-at-default instead of continuing
# to amortise. More realistic, but breaks exact reconciliation with the old CSV.
# Keep False for now so Phase 1 output matches your existing wide data.
FREEZE_TPOS_AT_DEFAULT = False

DB_PATH        = "ecl.db"
BASE_CSV       = "base_loans.csv"
PERF_CSV       = "performance_long.csv"

rng = np.random.default_rng(SEED)

# HELPERS
def months_between(a: date, b: date) -> int:
    return (b.year - a.year) * 12 + (b.month - a.month)

# 1. LOAN-LEVEL ATTRIBUTES  (feeds base_loans)
# ~82% as many customers as loans -> some customers hold multiple PLs.
n_customers   = int(N_LOANS * 0.82)
customer_ids  = rng.integers(0, n_customers, size=N_LOANS)

span          = (AS_OF - START).days
disb_dates    = [START + relativedelta(days=int(x)) for x in rng.integers(0, span, size=N_LOANS)]

amt_rupees    = np.clip(rng.lognormal(np.log(300_000), 0.6, N_LOANS), 50_000, 20_00_000)
amt_rupees    = np.round(amt_rupees, 2)                      # store paise precision
segment       = rng.choice(SEGMENTS, size=N_LOANS)
tenure        = rng.choice([12, 24, 36, 48, 60], N_LOANS, p=[0.10, 0.30, 0.30, 0.20, 0.10])

loan_ids      = np.array([f"LAN{1000000 + i}" for i in range(N_LOANS)])

base = pd.DataFrame({
    "distinct_loan_no": loan_ids,
    "customer_id":      [f"CUST{c:07d}" for c in customer_ids],
    "disbursal_date":   [d.isoformat() for d in disb_dates],
    "disbursal_amount": amt_rupees,          # RUPEES
    "segment":          segment.astype(int),
    "tenure_months":    tenure.astype(int),
})


# 2. PERFORMANCE TRAJECTORY  (feeds performance, long format)
age         = np.array([months_between(d, AS_OF) for d in disb_dates])          # months on book now
is_default  = rng.random(N_LOANS) < np.vectorize(SEG_DEFAULT_RATE.get)(segment)
default_mob = np.where(is_default, np.minimum(rng.integers(3, 40, N_LOANS), tenure), -1)
# outstanding-at-default (rupees): balance remaining when the loan first goes 90+
oad         = np.where(is_default,
                       np.round(amt_rupees * np.clip(1 - default_mob / tenure, 0, 1), 2),
                       0.0)

mob_arr = np.array(MOB_LIST)                                # shape (41,)
# amortising outstanding for every loan x MOB (rupees)
ratio   = np.clip(1 - mob_arr[None, :] / tenure[:, None], 0, 1)          # (N,41)
tpos    = np.round(amt_rupees[:, None] * ratio, 2)                        # (N,41)
# 90+ carried forward: once mob >= default_mob, exposure sits at OAD
defaulted_by = is_default[:, None] & (mob_arr[None, :] >= default_mob[:, None])
bad     = np.where(defaulted_by, oad[:, None], 0.0)                       # (N,41)

if FREEZE_TPOS_AT_DEFAULT:
    tpos = np.where(defaulted_by, oad[:, None], tpos)

# observed = cohort is old enough to have reached this MOB
observed = mob_arr[None, :] <= age[:, None]                              # (N,41) bool

# melt to long, keeping ONLY observed cells
r_idx, c_idx = np.where(observed)
perf = pd.DataFrame({
    "distinct_loan_no":      loan_ids[r_idx],
    "mob":                   mob_arr[c_idx],
    "tpos":                  tpos[r_idx, c_idx],
    "amt_90plus_settlement": bad[r_idx, c_idx],
})

# 3. WRITE CSVs
base.to_csv(BASE_CSV, index=False)
perf.to_csv(PERF_CSV, index=False)

# 4. LOAD INTO SQLITE  (with schema + indexes for fast joins)
con = sqlite3.connect(DB_PATH)
cur = con.cursor()
cur.executescript("""
DROP TABLE IF EXISTS base_loans;
DROP TABLE IF EXISTS performance;

CREATE TABLE base_loans (
    distinct_loan_no  TEXT    PRIMARY KEY,
    customer_id       TEXT    NOT NULL,
    disbursal_date    TEXT    NOT NULL,          -- ISO yyyy-mm-dd
    disbursal_amount  REAL    NOT NULL,          -- RUPEES
    segment           INTEGER NOT NULL,
    tenure_months     INTEGER NOT NULL
);

CREATE TABLE performance (
    distinct_loan_no       TEXT    NOT NULL,
    mob                    INTEGER NOT NULL,
    tpos                   REAL    NOT NULL,      -- RUPEES
    amt_90plus_settlement  REAL    NOT NULL,      -- RUPEES
    PRIMARY KEY (distinct_loan_no, mob)
);
""")

base.to_sql("base_loans", con, if_exists="append", index=False)
perf.to_sql("performance", con, if_exists="append", index=False)

cur.executescript("""
CREATE INDEX idx_perf_loan     ON performance (distinct_loan_no);
CREATE INDEX idx_perf_mob      ON performance (mob);
CREATE INDEX idx_base_segment  ON base_loans (segment);
""")
con.commit()

# 5. VALIDATION SUMMARY
print("=" * 60)
print("PHASE 0 COMPLETE")
print("=" * 60)
print(f"AS_OF date            : {AS_OF}")
print(f"base_loans rows       : {len(base):,}")
print(f"performance rows      : {len(perf):,}")
print(f"distinct customers    : {base.customer_id.nunique():,}")
print(f"max loans / customer  : {base.groupby('customer_id').size().max()}")
print(f"disbursal span        : {base.disbursal_date.min()} -> {base.disbursal_date.max()}")
print(f"total disbursal (cr)  : {base.disbursal_amount.sum() / 1e7:,.2f}")
print(f"avg observed MOB/loan : {len(perf) / len(base):.1f}")

# triangle sanity: an old loan has many MOB rows, a new loan has few
tri = perf.groupby('distinct_loan_no').mob.max().describe()[['min', '50%', 'max']]
print(f"max-observed-MOB (min/med/max): {tri['min']:.0f} / {tri['50%']:.0f} / {tri['max']:.0f}")

# default rate by segment (should rise with segment)
seg_def = (perf.groupby('distinct_loan_no').amt_90plus_settlement.max()
           .gt(0).groupby(base.set_index('distinct_loan_no').segment).mean())
print("default rate by segment:")
for s, v in seg_def.items():
    print(f"    segment {s}: {v:.1%}")

# round-trip check against DB
n_base = cur.execute("SELECT COUNT(*) FROM base_loans").fetchone()[0]
n_perf = cur.execute("SELECT COUNT(*) FROM performance").fetchone()[0]
print(f"DB check              : base_loans={n_base:,}  performance={n_perf:,}")
con.close()
print(f"\nWrote: {DB_PATH}, {BASE_CSV}, {PERF_CSV}")