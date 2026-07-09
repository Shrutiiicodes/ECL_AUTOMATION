"""
PHASE 3 - TRIANGLE + CHAIN-LADDER ENGINE  (modelling core)
==========================================================
Turns the DATA_ECL_NEW summary into completed run-off triangles for 90+ and
TPOS, filling the immature (yellow) cells with the bank's exact Excel formula.

Triangle layout (mirrors the Workings sheet)
    rows    = FY_QUARTER cohorts (oldest at top)
    columns = MOB (0,3,...,120)
    cell    = rate = amount(cohort, MOB) / DISBURSAL_AMT(cohort)

Maturity (replaces manual yellow-highlighting)
    A quarter-cohort is MATURE at MOB j only if its whole disbursal window has
    had j months to develop: months_between(quarter_end, AS_OF) >= j.
    Cells beyond that are immature and get projected.

Chain-ladder fill (exact reproduction of the Excel formula)
    For an immature cell at (row R, MOB X):
        value(R,X) = value(R-1,X)
                     * SUMPRODUCT(X[top..R-1], DISB[top..R-1])
                     / SUMPRODUCT(X[top..R-2], DISB[top..R-2])
    i.e. carry the cohort above down the SAME column, scaled by the
    disbursal-weighted cumulative trend. IFERROR -> 0 when the denominator is 0.
    Cells fill top-to-bottom so each projection can feed the next (as in Excel).

Outputs: tri_90plus_rate.csv, tri_tpos_rate.csv,
         tri_90plus_amt.csv,  tri_tpos_amt.csv,
         triangle_90plus_highlighted.xlsx  (projected cells shaded yellow)
"""

import calendar
from datetime import date

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# CONFIG - all knobs live in config.py
from config import *      # FEED_CSV, AS_OF, MOB_LIST, SEGMENT

# PHASE 2 (folded in): load the summary and slice to a triangle-ready frame
def load_summary(path=FEED_CSV, segment=None):
    df = pd.read_csv(path)
    if segment is not None:
        df = df[df.SEGMENT == segment].copy()
    # collapse segments -> one row per FY_QUARTER (sum amounts + disbursal + count)
    amt_cols = [c for c in df.columns if c.startswith(("AMT_90PLUS", "TPOS_"))]
    agg = {"LAN_CNT": "sum", "DISBURSAL_AMT": "sum", **{c: "sum" for c in amt_cols}}
    g = df.groupby("FY_QUARTER").agg(agg)
    key = lambda q: (int(q[2:4]), int(q[-1]))
    return g.reindex(sorted(g.index, key=key))

# MATURITY
def quarter_end(label):                      # 'FY16-Q1' -> 2015-06-30
    fy = 2000 + int(label[2:4]); q = int(label[-1])
    if   q == 1: y, m = fy - 1, 6
    elif q == 2: y, m = fy - 1, 9
    elif q == 3: y, m = fy - 1, 12
    else:        y, m = fy,     3
    return date(y, m, calendar.monthrange(y, m)[1])

def max_mature_mob(label, as_of):
    qe = quarter_end(label)
    months = (as_of.year - qe.year) * 12 + (as_of.month - qe.month)
    valid = [m for m in MOB_LIST if m <= months]
    return max(valid) if valid else -1

# CHAIN-LADDER FILL  (exact Excel-formula reproduction, down each column)
def chain_ladder_fill(rate, disb, mature):
    """rate, mature: DataFrames [quarter x MOB]; disb: Series[quarter]. Returns filled rate."""
    R = rate.to_numpy(dtype=float).copy()
    M = mature.to_numpy()
    w = disb.to_numpy(dtype=float)
    n_rows, n_cols = R.shape
    for cj in range(n_cols):
        for ri in range(n_rows):
            if M[ri, cj]:
                continue                                     # observed -> keep
            if ri == 0:
                R[ri, cj] = 0.0                              # no cohort above
                continue
            prev = R[ri - 1, cj]
            num = np.nansum(R[:ri,     cj] * w[:ri])          # rows top..R-1
            den = np.nansum(R[:ri - 1, cj] * w[:ri - 1])      # rows top..R-2
            R[ri, cj] = prev * num / den if den != 0 else 0.0
    return pd.DataFrame(R, index=rate.index, columns=rate.columns)

# BUILD TRIANGLES
def build(metric_prefix, feed):
    cols = [f"{metric_prefix}{m}MOB" for m in MOB_LIST]
    amt  = feed[cols].copy(); amt.columns = MOB_LIST
    disb = feed["DISBURSAL_AMT"]
    rate = amt.div(disb, axis=0)                             # rate = amount / disbursal
    mature = pd.DataFrame(
        {m: [MOB_LIST[i] <= max_mature_mob(q, AS_OF) and m <= max_mature_mob(q, AS_OF)
             for q in feed.index] for i, m in enumerate(MOB_LIST)},
        index=feed.index)
    mature = pd.DataFrame(
        [[m <= max_mature_mob(q, AS_OF) for m in MOB_LIST] for q in feed.index],
        index=feed.index, columns=MOB_LIST)
    rate_masked = rate.where(mature)                         # immature -> NaN
    rate_full   = chain_ladder_fill(rate_masked, disb, mature)
    amt_full    = rate_full.mul(disb, axis=0)               # back to crores
    return rate_full, amt_full, mature, disb


feed = load_summary(FEED_CSV, SEGMENT)
r90, a90, mat90, disb = build("AMT_90PLUS_SETTLEMENT_", feed)
rtp, atp, mattp, _    = build("TPOS_", feed)

r90.to_csv("tri_90plus_rate.csv"); a90.to_csv("tri_90plus_amt.csv")
rtp.to_csv("tri_tpos_rate.csv");   atp.to_csv("tri_tpos_amt.csv")

# HIGHLIGHTED XLSX  (visual check: yellow = projected)
def highlighted_xlsx(rate_full, mature, disb, path, title):
    wb = Workbook(); ws = wb.active; ws.title = title
    HF = PatternFill("solid", fgColor="1F4E78"); HFONT = Font(bold=True, color="FFFFFF")
    YEL = PatternFill("solid", fgColor="FFFF00"); C = Alignment("center", "center")
    BD = Border(*[Side(style="thin", color="D9D9D9")] * 4)
    heads = ["FY_QUARTER", "DISB_AMT"] + [f"{m}MOB" for m in MOB_LIST]
    for j, h in enumerate(heads, 1):
        c = ws.cell(1, j, h); c.fill, c.font, c.alignment, c.border = HF, HFONT, C, BD
    for i, q in enumerate(rate_full.index):
        r = 2 + i
        ws.cell(r, 1, q).border = BD
        dc = ws.cell(r, 2, round(float(disb.iloc[i]), 4)); dc.number_format = "#,##0.0000"; dc.border = BD
        for j, m in enumerate(MOB_LIST):
            cell = ws.cell(r, 3 + j, round(float(rate_full.iloc[i, j]), 6))
            cell.number_format = "0.00%"; cell.alignment = C; cell.border = BD
            if not mature.iloc[i, j]:
                cell.fill = YEL                              # projected cell
    ws.freeze_panes = "C2"; ws.column_dimensions["A"].width = 11
    wb.save(path)

highlighted_xlsx(r90, mat90, disb, "triangle_90plus_highlighted.xlsx", "Tri_90plus")

# VALIDATION
proj = (~mat90).sum().sum()
tot  = mat90.size
print("=" * 60)
print(f"PHASE 3 COMPLETE   (segment = {SEGMENT or 'ALL'})")
print("=" * 60)
print(f"triangle shape        : {r90.shape[0]} cohorts x {r90.shape[1]} MOB")
print(f"cells                 : {tot}  |  observed {tot-proj}  |  projected {proj}")

print("\nmax mature MOB per cohort (first 3, last 3):")
labels = list(feed.index)
for q in labels[:3] + labels[-3:]:
    print(f"    {q}: {max_mature_mob(q, AS_OF)}")

# worked example: recompute one projected cell by hand and compare to engine
ri = next(i for i in range(len(labels)) if not mat90.iloc[i].all())   # first immature cohort
cj = next(j for j in range(len(MOB_LIST)) if not mat90.iloc[ri, j])   # its first immature MOB
w  = disb.to_numpy()
num = np.nansum(r90.iloc[:ri, cj].to_numpy() * w[:ri])
den = np.nansum(r90.iloc[:ri-1, cj].to_numpy() * w[:ri-1])
hand = r90.iloc[ri-1, cj] * num / den if den else 0.0
print(f"\nworked example  cell [{labels[ri]}, {MOB_LIST[cj]}MOB]:")
print(f"    prev(above)={r90.iloc[ri-1,cj]:.6f}  factor={num/den if den else 0:.6f}")
print(f"    hand={hand:.8f}   engine={r90.iloc[ri,cj]:.8f}   match={np.isclose(hand, r90.iloc[ri,cj])}")

# sanity: no NaN left, 90+ rate non-decreasing across MOB (cumulative defaults)
print(f"\nNaN remaining         : {int(r90.isna().sum().sum())}  (should be 0)")
nondec = (r90.diff(axis=1).fillna(0) >= -1e-9).all().all()
print(f"90+ rate non-decreasing across MOB : {nondec}")
print(f"\nWrote: tri_90plus_rate.csv, tri_tpos_rate.csv, tri_90plus_amt.csv,")
print(f"       tri_tpos_amt.csv, triangle_90plus_highlighted.xlsx")