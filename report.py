"""
EXCEL REPORT GENERATION
=======================
Assembles all upstream outputs into ONE formatted workbook (ECL_Report.xlsx).
No recomputation - it only lays out DataFrames the pipeline already produced.

Tabs
  Summary          cover metrics + headline weighted loss rate
  DATA_ECL_NEW     the SQL feed (FY_QUARTER x SEGMENT)
  Pivot_90plus     90+ amount triangle, yellow = chain-ladder projected
  Pivot_TPOS       TPOS amount triangle, yellow = projected
  BadRate_90plus   90+ / DISB rate triangle (PD curve), yellow = projected
  Movements        TPOS + 90+ movement tables (12..120)
  LossRate_Qtr     per-quarter loss rate at 84M and 120M (>100% flagged red)
  Weighted_LR      SUMPRODUCT(LR, DISB)/SUM(DISB) per observation window

-------------------------------------------------------------------------------
This module exposes ONE function:

    build_excel(feed, tris, lrr, ecl, path=OUT) -> Workbook

It takes the upstream results IN MEMORY:
    feed  : the SQL summary DataFrame            (sql_refactor.run().feed)
    tris  : chain_ladder.run() result            (.a90, .atp, .r90)
    lrr   : loss_rate.run() result               (.mv90, .mvtp)
    ecl   : final_ecl.run() result               (.by_quarter, .wavg)
Writing the workbook IS this module's job, so build_excel does save to `path`.
The standalone `__main__` block reconstructs those inputs from the CSVs on disk.
"""

import calendar
from datetime import date

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from config import *      # AS_OF, MOB_LIST, ANCHORS, OUT

# --------------------------------------------------------------------------- #
# presentation constants (styles are presentation, not configuration)
# --------------------------------------------------------------------------- #
HF   = PatternFill("solid", fgColor="1F4E78"); HFONT = Font(bold=True, color="FFFFFF", size=10)
IFL  = PatternFill("solid", fgColor="DDEBF7"); IFONT = Font(bold=True, size=10)
YEL  = PatternFill("solid", fgColor="FFFF00"); WARN = PatternFill("solid", fgColor="FFC7CE")
TOT  = PatternFill("solid", fgColor="C6E0B4")
TITLE = Font(bold=True, size=14, color="1F4E78"); SUB = Font(italic=True, size=9, color="808080")
CF = Font(size=10); C = Alignment("center", "center"); L = Alignment("left", "center")
BD = Border(*[Side(style="thin", color="D9D9D9")] * 4)
CR, PC, IN = "#,##0.0000", "0.00%", "#,##0"


def quarter_end(label):
    fy = 2000 + int(label[2:4]); q = int(label[-1])
    y, m = {1: (fy - 1, 6), 2: (fy - 1, 9), 3: (fy - 1, 12), 4: (fy, 3)}[q]
    return date(y, m, calendar.monthrange(y, m)[1])


def max_mature_mob(label):
    qe = quarter_end(label)
    months = (AS_OF.year - qe.year) * 12 + (AS_OF.month - qe.month)
    v = [m for m in MOB_LIST if m <= months]
    return max(v) if v else -1


is_mature = lambda q, mob: mob <= max_mature_mob(q)


def write_triangle(ws, matrix, fmt, prefix=""):
    hr = 1
    for j, h in enumerate(["FY_QUARTER"] + [f"{prefix}{m}MOB" for m in MOB_LIST], 1):
        c = ws.cell(hr, j, h); c.fill, c.font, c.alignment, c.border = HF, HFONT, C, BD
    for i, q in enumerate(matrix.index):
        r = hr + 1 + i
        ic = ws.cell(r, 1, q); ic.fill, ic.font, ic.alignment, ic.border = IFL, IFONT, C, BD
        for j, m in enumerate(MOB_LIST):
            c = ws.cell(r, 2 + j, round(float(matrix.iloc[i, j]), 6))
            c.number_format, c.font, c.alignment, c.border = fmt, CF, C, BD
            if not is_mature(q, m): c.fill = YEL
    ws.column_dimensions["A"].width = 12
    for j in range(2, 2 + len(MOB_LIST)): ws.column_dimensions[get_column_letter(j)].width = 11
    ws.freeze_panes = ws.cell(hr + 1, 2)


def mv_block(ws, r0, prefix, frame):
    for j, h in enumerate(["FY_QUARTER"] + [f"{prefix}{m}MOB" for m in ANCHORS], 1):
        c = ws.cell(r0, j, h); c.fill, c.font, c.alignment, c.border = HF, HFONT, C, BD
    for i, q in enumerate(frame.index):
        rr = r0 + 1 + i
        ic = ws.cell(rr, 1, q); ic.fill, ic.font, ic.border = IFL, IFONT, BD
        for j, m in enumerate(ANCHORS):
            c = ws.cell(rr, 2 + j, round(float(frame.loc[q, m]), 6))
            c.number_format, c.alignment, c.border = CR, C, BD
    return r0 + 1 + len(frame)


def build_excel(feed, tris, lrr, ecl, path=OUT):
    """Lay out every upstream DataFrame into the formatted workbook and save it."""
    a90, atp, r90 = tris.a90.copy(), tris.atp.copy(), tris.r90.copy()
    mv90, mvtp = lrr.mv90.copy(), lrr.mvtp.copy()
    qtr, wavg = ecl.by_quarter, ecl.wavg
    for df in (a90, atp, r90, mv90, mvtp):
        df.columns = [int(c) for c in df.columns]   # robust to int- or str-keyed columns

    wb = Workbook()

    # 1) Summary ---------------------------------------------------------------
    ws = wb.active; ws.title = "Summary"
    hl = wavg.iloc[0]
    rows = [
        ("As-of date", str(AS_OF)),
        ("Cohorts (FY quarters)", len(a90)),
        ("MOB grid", f"0..120 step 3  ({len(MOB_LIST)} points)"),
        ("Loss-rate anchors", "84M and 120M"),
        ("Headline window", hl.WINDOW),
        ("Weighted-avg loss rate", hl.WEIGHTED_AVG_LR),
        ("Portfolio current TPOS (cr)", round(qtr.CURRENT_TPOS.sum(), 2)),
        ("Portfolio ECL (cr) [provisional]", round(hl.WEIGHTED_AVG_LR * qtr.CURRENT_TPOS.sum(), 2)),
        ("Final-ECL rule", "weighted-avg LR x current TPOS (multiplier unconfirmed)"),
    ]
    for i, (k, v) in enumerate(rows):
        r = 1 + i
        kc = ws.cell(r, 1, k); kc.font, kc.fill, kc.border, kc.alignment = IFONT, IFL, BD, L
        vc = ws.cell(r, 2, v); vc.border, vc.alignment = BD, L
        if k == "Weighted-avg loss rate": vc.number_format = PC
    ws.column_dimensions["A"].width = 32; ws.column_dimensions["B"].width = 46

    # 2) DATA_ECL_NEW ----------------------------------------------------------
    ws = wb.create_sheet("DATA_ECL_NEW")
    for j, h in enumerate(feed.columns, 1):
        c = ws.cell(1, j, h); c.fill, c.font, c.alignment, c.border = HF, HFONT, C, BD
    for i, row in feed.iterrows():
        r = 2 + i
        for j, col in enumerate(feed.columns, 1):
            c = ws.cell(r, j, row[col]); c.font, c.border = CF, BD
            if col == "LAN_CNT": c.number_format = IN
            elif col not in ("FY_QUARTER", "SEGMENT"): c.number_format = CR
    ws.freeze_panes = "C2"; ws.column_dimensions["A"].width = 12

    # 3-5) triangles -----------------------------------------------------------
    write_triangle(wb.create_sheet("Pivot_90plus"), a90, CR)
    write_triangle(wb.create_sheet("Pivot_TPOS"), atp, CR)
    write_triangle(wb.create_sheet("BadRate_90plus"), r90, PC, "PD_")

    # 6) Movements -------------------------------------------------------------
    ws = wb.create_sheet("Movements")
    end = mv_block(ws, 1, "TPOS_", mvtp)
    mv_block(ws, end + 2, "90PLUS_", mv90)
    ws.column_dimensions["A"].width = 12

    # 7) LossRate_Qtr ----------------------------------------------------------
    ws = wb.create_sheet("LossRate_Qtr")
    for j, h in enumerate(["FY_QUARTER", "DISB_AMT (weight)", "LOSS_RATE_84M", "LOSS_RATE_120M",
                           "CURRENT_MOB", "CURRENT_TPOS"], 1):
        c = ws.cell(1, j, h); c.fill, c.font, c.alignment, c.border = HF, HFONT, C, BD
    for i, r in qtr.iterrows():
        rr = 2 + i
        ic = ws.cell(rr, 1, r.FY_QUARTER); ic.fill, ic.font = IFL, IFONT
        ws.cell(rr, 2, round(r.DISBURSAL_AMT, 4)).number_format = CR
        c84 = ws.cell(rr, 3, round(r.LOSS_RATE_84M, 6)); c84.number_format = PC
        c120 = ws.cell(rr, 4, round(r.LOSS_RATE_120M, 6)); c120.number_format = PC
        if r.LOSS_RATE_120M > 1: c120.fill = WARN
        if r.LOSS_RATE_84M > 1: c84.fill = WARN
        ws.cell(rr, 5, int(r.CURRENT_MOB))
        ws.cell(rr, 6, round(r.CURRENT_TPOS, 6)).number_format = CR
        for cc in range(1, 7): ws.cell(rr, cc).alignment = C; ws.cell(rr, cc).border = BD
    ws.column_dimensions["A"].width = 12
    for col in "BCDEF": ws.column_dimensions[col].width = 16

    # 8) Weighted_LR -----------------------------------------------------------
    ws = wb.create_sheet("Weighted_LR")
    heads = ["WINDOW", "FY_START", "FY_END", "ANCHOR", "N_QTRS", "TOTAL_DISB", "SIMPLE_AVG", "WEIGHTED_AVG"]
    for j, h in enumerate(heads, 1):
        c = ws.cell(1, j, h); c.fill, c.font, c.alignment, c.border = HF, HFONT, C, BD
    for i, r in wavg.iterrows():
        rr = 2 + i
        ws.cell(rr, 1, r.WINDOW); ws.cell(rr, 2, r.FY_START); ws.cell(rr, 3, r.FY_END)
        ws.cell(rr, 4, int(r.ANCHOR_MOB)); ws.cell(rr, 5, int(r.N_QUARTERS))
        ws.cell(rr, 6, round(r.TOTAL_DISB, 4)).number_format = CR
        ws.cell(rr, 7, round(r.SIMPLE_AVG_LR, 6)).number_format = PC
        wc = ws.cell(rr, 8, round(r.WEIGHTED_AVG_LR, 6)); wc.number_format = PC; wc.font = Font(bold=True)
        if r.WEIGHTED_AVG_LR > 1: wc.fill = WARN
        for cc in range(1, 9): ws.cell(rr, cc).alignment = C; ws.cell(rr, cc).border = BD
    r0 = 3 + len(wavg)
    ws.cell(r0, 1, "Portfolio current TPOS (cr)").font = Font(bold=True)
    ws.cell(r0, 2, round(qtr.CURRENT_TPOS.sum(), 4)).number_format = CR
    ws.cell(r0 + 1, 1, "Portfolio ECL (cr) [provisional]").font = Font(bold=True)
    ec = ws.cell(r0 + 1, 2, round(hl.WEIGHTED_AVG_LR * qtr.CURRENT_TPOS.sum(), 4))
    ec.number_format = CR; ec.fill = TOT
    for col, w in zip("ABCDEFGH", [20, 10, 10, 9, 9, 14, 13, 14]):
        ws.column_dimensions[col].width = w

    wb.save(path)
    return wb


# =========================================================================== #
# Standalone entrypoint: reconstruct the upstream inputs from the CSVs on disk.
# None of this runs on import.
# =========================================================================== #
if __name__ == "__main__":
    from types import SimpleNamespace

    feed = pd.read_csv("DATA_ECL_NEW.csv")
    a90 = pd.read_csv("tri_90plus_amt.csv", index_col=0);  a90.columns = [int(c) for c in a90.columns]
    atp = pd.read_csv("tri_tpos_amt.csv",  index_col=0);  atp.columns = [int(c) for c in atp.columns]
    r90 = pd.read_csv("tri_90plus_rate.csv", index_col=0); r90.columns = [int(c) for c in r90.columns]
    mv90 = pd.read_csv("movement_90plus.csv", index_col=0); mv90.columns = [int(c) for c in mv90.columns]
    mvtp = pd.read_csv("movement_tpos.csv",  index_col=0); mvtp.columns = [int(c) for c in mvtp.columns]
    qtr  = pd.read_csv("ecl_by_quarter.csv")
    wavg = pd.read_csv("weighted_loss_rate.csv")

    tris = SimpleNamespace(a90=a90, atp=atp, r90=r90)
    lrr  = SimpleNamespace(mv90=mv90, mvtp=mvtp)
    ecl  = SimpleNamespace(by_quarter=qtr, wavg=wavg)

    wb = build_excel(feed, tris, lrr, ecl, OUT)

    hl = wavg.iloc[0]
    print("=" * 60); print("REPORT COMPLETE"); print("=" * 60)
    print(f"workbook : {OUT}")
    print(f"sheets   : {wb.sheetnames}")
    print(f"headline : {hl.WINDOW}  weighted LR = {hl.WEIGHTED_AVG_LR:.4%}")