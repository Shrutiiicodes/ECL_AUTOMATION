"""
PHASE 1 - SQL REFACTOR  (the headline deliverable)
==================================================
Replaces the bank's ~82 repetitive LEFT JOINs (one per MOB per metric) with a
SINGLE aggregation, and *generates* that SQL from the MOB list so the query is
never hand-edited. Produces DATA_ECL_NEW: one row per FY_QUARTER x SEGMENT.

What the generated SQL does, in order
-------------------------------------
1. perf_wide : collapse the long `performance` table to ONE row per loan using
               conditional aggregation  SUM(CASE WHEN mob=X THEN metric END).
               This one CTE is the equivalent of all 82 bank joins.
2. base_fy   : take base_loans, filter to the disbursal window, and DERIVE
               fy_quarter from disbursal_date (Indian FY, 'FY16-Q1' style) --
               the SQL analog of the bank's concat/case expression.
3. final     : join (1:1) and group by FY_QUARTER x SEGMENT ->
               LAN_CNT, DISBURSAL_AMT, 90+ per MOB, TPOS per MOB,
               all divided by 1e7 to report in CRORES (the bank's /10^7).

Why aggregate-before-join: joining the long performance table directly would
fan each loan into ~23 rows and inflate COUNT/SUM. Collapsing to perf_wide first
keeps the base<->performance join 1:1.

Parameters (edit CONFIG): disbursal window + MOB grid. as-of maturity is handled
later in Phase 3, not here.

-------------------------------------------------------------------------------
This module exposes a PURE function:

    run(db_path, start_disb, end_disb, mob_list) -> SqlOutput(feed, sql)

It connects to the database, runs the generated query, and returns the summary
DataFrame IN MEMORY (plus the generated SQL text for review). It writes nothing.
The CSV/.sql side effects and the independent reconciliation live only in the
`if __name__ == "__main__"` block below.

Real-data cutover: point `db_path` at the bank warehouse and adjust the dialect
in build_summary_sql / FY_QUARTER_SQL. Nothing downstream changes.
"""

import sqlite3
from typing import NamedTuple

import pandas as pd

from config import *      # DB_PATH, MOB_LIST, START_DISB, END_DISB, OUT_CSV, OUT_SQL


class SqlOutput(NamedTuple):
    feed: pd.DataFrame   # DATA_ECL_NEW: one row per FY_QUARTER x SEGMENT
    sql: str             # the generated query, for review / versioning


# FY-QUARTER EXPRESSION  (Indian FY, matches data_generation.py exactly)
# Apr-Jun=Q1, Jul-Sep=Q2, Oct-Dec=Q3, Jan-Mar=Q4.  FY label = ending year.
FY_QUARTER_SQL = """'FY' ||
    substr(CAST(CASE WHEN CAST(strftime('%m', disbursal_date) AS INT) IN (1,2,3)
                     THEN strftime('%Y', disbursal_date)
                     ELSE strftime('%Y', disbursal_date) + 1 END AS TEXT), 3, 2)
    || '-Q' ||
    CASE WHEN CAST(strftime('%m', disbursal_date) AS INT) IN (4,5,6)   THEN 1
         WHEN CAST(strftime('%m', disbursal_date) AS INT) IN (7,8,9)   THEN 2
         WHEN CAST(strftime('%m', disbursal_date) AS INT) IN (10,11,12) THEN 3
         ELSE 4 END"""

# sort key so FY16-Q1 < FY16-Q2 < ... < FY26-Q2
FY_SORT_SQL = "substr(fy_quarter,3,2), substr(fy_quarter,7,1)"


# --------------------------------------------------------------------------- #
# SQL GENERATOR  (this is what replaces the 82 hand-written joins)
# --------------------------------------------------------------------------- #
def build_summary_sql(mob_list):
    ind = "        "
    perf_wide_cols = ",\n".join(
        f"{ind}SUM(CASE WHEN mob={m} THEN amt_90plus_settlement END) AS bad_{m},\n"
        f"{ind}SUM(CASE WHEN mob={m} THEN tpos END)                  AS tpos_{m}"
        for m in mob_list
    )
    sel_bad  = ",\n".join(f"    ROUND(SUM(pw.bad_{m})/1e7, 6)  AS AMT_90PLUS_SETTLEMENT_{m}MOB" for m in mob_list)
    sel_tpos = ",\n".join(f"    ROUND(SUM(pw.tpos_{m})/1e7, 6) AS TPOS_{m}MOB" for m in mob_list)

    return f"""
WITH perf_wide AS (          -- 1 row per loan; this CTE == all 82 bank joins
    SELECT distinct_loan_no,
{perf_wide_cols}
    FROM performance
    GROUP BY distinct_loan_no
),
base_fy AS (                 -- filter window + derive FY quarter in SQL
    SELECT distinct_loan_no, segment, disbursal_amount,
           {FY_QUARTER_SQL} AS fy_quarter
    FROM base_loans
    WHERE disbursal_date >= ? AND disbursal_date <= ?
)
SELECT
    b.fy_quarter                       AS FY_QUARTER,
    b.segment                          AS SEGMENT,
    COUNT(*)                           AS LAN_CNT,
    ROUND(SUM(b.disbursal_amount)/1e7, 6) AS DISBURSAL_AMT,
{sel_bad},
{sel_tpos}
FROM base_fy b
LEFT JOIN perf_wide pw ON b.distinct_loan_no = pw.distinct_loan_no
GROUP BY b.fy_quarter, b.segment
ORDER BY {FY_SORT_SQL}, b.segment;
""".strip()


def run(db_path=DB_PATH, start_disb=START_DISB, end_disb=END_DISB, mob_list=MOB_LIST) -> SqlOutput:
    """Generate the query, run it against the DB, return the summary feed. No file I/O."""
    sql = build_summary_sql(mob_list)
    con = sqlite3.connect(db_path)
    try:
        feed = pd.read_sql(sql, con, params=[start_disb, end_disb])
    finally:
        con.close()
    return SqlOutput(feed=feed, sql=sql)


# =========================================================================== #
# Standalone entrypoint: write DATA_ECL_NEW.csv + the generated .sql, then run
# an INDEPENDENT pandas reconciliation. None of this runs on import.
# =========================================================================== #
def _reconcile(feed: pd.DataFrame, db_path=DB_PATH) -> None:
    """Recompute a spot-check via a plain pandas groupby and diff against the SQL feed."""
    con = sqlite3.connect(db_path)
    base = pd.read_sql("SELECT * FROM base_loans", con)
    perf = pd.read_sql("SELECT * FROM performance", con)
    con.close()

    def fy_q(d):
        d = pd.Timestamp(d); m = d.month
        fy = d.year if m in (1, 2, 3) else d.year + 1
        q = 4 if m in (1, 2, 3) else 1 if m in (4, 5, 6) else 2 if m in (7, 8, 9) else 3
        return f"FY{str(fy)[-2:]}-Q{q}"

    base = base[(base.disbursal_date >= START_DISB) & (base.disbursal_date <= END_DISB)].copy()
    base["fy_quarter"] = base.disbursal_date.map(fy_q)
    bad12 = (perf[perf.mob == 12].groupby("distinct_loan_no").amt_90plus_settlement.sum()
             .rename("bad12"))                                  # spot-check one MOB
    merged = base.set_index("distinct_loan_no").join(bad12)
    g = merged.groupby(["fy_quarter", "segment"])
    exp_cnt  = g.size()
    exp_disb = g["disbursal_amount"].sum() / 1e7
    exp_bad0 = g["bad12"].sum() / 1e7

    sql_idx = feed.set_index(["FY_QUARTER", "SEGMENT"])
    print("=" * 60)
    print("PHASE 1 COMPLETE  -  reconciliation vs independent pandas")
    print("=" * 60)
    print(f"SQL replaced          : {len(MOB_LIST)*2} joins  ->  1 aggregation CTE")
    print(f"summary rows          : {len(feed)}  (FY_QUARTER x SEGMENT)")
    print(f"columns               : {feed.shape[1]}  (2 keys + LAN_CNT + DISB + {len(MOB_LIST)}x2 MOB)")
    print(f"LAN_CNT total         : {feed.LAN_CNT.sum():,}  (should be <= 60,000)")
    print(f"DISBURSAL_AMT total   : {feed.DISBURSAL_AMT.sum():,.2f} cr")

    cnt_ok  = (sql_idx.LAN_CNT.reindex(exp_cnt.index) == exp_cnt).all()
    disb_ok = (sql_idx.DISBURSAL_AMT.reindex(exp_disb.index) - exp_disb).abs().max()
    bad_ok  = (sql_idx.AMT_90PLUS_SETTLEMENT_12MOB.reindex(exp_bad0.index).fillna(0) - exp_bad0.fillna(0)).abs().max()
    print(f"LAN_CNT matches       : {cnt_ok}")
    print(f"DISBURSAL max diff    : {disb_ok:.2e} cr")
    print(f"90+ @12MOB max diff   : {bad_ok:.2e} cr")


if __name__ == "__main__":
    out = run()

    with open(OUT_SQL, "w") as f:
        f.write(out.sql)
    out.feed.to_csv(OUT_CSV, index=False)

    _reconcile(out.feed, DB_PATH)
    print(f"\nWrote: {OUT_CSV}, {OUT_SQL}")
    print("\nfirst rows:")
    print(out.feed.iloc[:6, :6].to_string(index=False))