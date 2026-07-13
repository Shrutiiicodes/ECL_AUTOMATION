"""
EXCEL REPORT GENERATION  (live-formula edition)
===============================================
Source sheets carry VALUES; every COMPUTED sheet carries live Excel FORMULAS
referencing the source cells, so a reviewer can click any PD / loss-rate /
weighted-average cell and trace it to the amount triangle - like the manual
workbook this pipeline replaces.

  VALUES (inputs / model outputs, yellow when projected):
      DATA_ECL_NEW, Pivot_90plus, Pivot_TPOS   (col B = DISB)
  FORMULAS (computed in-cell):
      BadRate_90plus = 90+/DISB ; Movements = links at anchor MOBs ;
      LossRate_Qtr = 90+@A / SUM(TPOS 12..A) ;
      Weighted_LR = SUMPRODUCT(LR,DISB)/SUM(DISB) ; Summary links to Weighted_LR

Python still computes everything (validation.py reconciles independently); the
formulas recompute to the identical number. fullCalcOnLoad makes the workbook
show results the moment it opens.

    build_excel(feed, tris, lrr, ecl, path=OUT) -> Workbook
"""

import calendar
from datetime import date

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from config import *      # AS_OF, MOB_LIST, ANCHORS, ANCHOR_MOBS, anchors_for, WINDOWS, HEADLINE, fy_key, OUT

HF   = PatternFill("solid", fgColor="1F4E78"); HFONT = Font(bold=True, color="FFFFFF", size=10)
IFL  = PatternFill("solid", fgColor="DDEBF7"); IFONT = Font(bold=True, size=10)
YEL  = PatternFill("solid", fgColor="FFFF00"); WARN = PatternFill("solid", fgColor="FFC7CE")
TOT  = PatternFill("solid", fgColor="C6E0B4")
CF = Font(size=10); C = Alignment("center", "center"); L = Alignment("left", "center")
BD = Border(*[Side(style="thin", color="D9D9D9")] * 4)
CR, PC, IN = "#,##0.0000", "0.00%", "#,##0"

MOB_COL0 = 3                                        # column index of MOB_LIST[0] (A=FY, B=DISB, C=MOB0)
def mob_col(j):        return MOB_COL0 + j
def mob_letter(j):     return get_column_letter(mob_col(j))
def anchor_letter(a):  return mob_letter(MOB_LIST.index(a))


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


def _triangle_header(ws, prefix):
    for j, h in enumerate(["FY_QUARTER", "DISB_AMT"] + [f"{prefix}{m}MOB" for m in MOB_LIST], 1):
        c = ws.cell(1, j, h); c.fill, c.font, c.alignment, c.border = HF, HFONT, C, BD
    ws.column_dimensions["A"].width = 12; ws.column_dimensions["B"].width = 12
    for j in range(MOB_COL0, MOB_COL0 + len(MOB_LIST)):
        ws.column_dimensions[get_column_letter(j)].width = 11
    ws.freeze_panes = ws.cell(2, MOB_COL0)


def write_amount_triangle(ws, amt, disb, prefix=""):
    """Source amount triangle: DISB + amounts as VALUES, yellow = projected."""
    _triangle_header(ws, prefix)
    for i, q in enumerate(amt.index):
        r = 2 + i
        ic = ws.cell(r, 1, q); ic.fill, ic.font, ic.alignment, ic.border = IFL, IFONT, C, BD
        dc = ws.cell(r, 2, round(float(disb.iloc[i]), 6)); dc.number_format, dc.border, dc.alignment = CR, BD, C
        for j, m in enumerate(MOB_LIST):
            c = ws.cell(r, mob_col(j), round(float(amt.iloc[i, j]), 6))
            c.number_format, c.font, c.alignment, c.border = CR, CF, C, BD
            if not is_mature(q, m): c.fill = YEL


def write_badrate_triangle(ws, index, src_sheet="Pivot_90plus", prefix="PD_"):
    """PD curve as FORMULAS: = 90+ amount / disbursal, referencing src_sheet."""
    _triangle_header(ws, prefix)
    for i, q in enumerate(index):
        r = 2 + i
        ic = ws.cell(r, 1, q); ic.fill, ic.font, ic.alignment, ic.border = IFL, IFONT, C, BD
        dc = ws.cell(r, 2, f"='{src_sheet}'!$B{r}"); dc.number_format, dc.border, dc.alignment = CR, BD, C
        for j, m in enumerate(MOB_LIST):
            c = ws.cell(r, mob_col(j), f"=IFERROR('{src_sheet}'!{mob_letter(j)}{r}/'{src_sheet}'!$B{r},0)")
            c.number_format, c.font, c.alignment, c.border = PC, CF, C, BD
            if not is_mature(q, m): c.fill = YEL


def build_excel(feed, tris, lrr, ecl, path=OUT):
    a90, atp = tris.a90.copy(), tris.atp.copy()
    disb = tris.disb.copy()
    qtr, wavg = ecl.by_quarter, ecl.wavg
    for df in (a90, atp):
        df.columns = [int(c) for c in df.columns]
    cohorts = list(a90.index)
    row_of = {q: 2 + i for i, q in enumerate(cohorts)}

    wb = Workbook()
    wb.calculation.fullCalcOnLoad = True

    # 1) Summary ---------------------------------------------------------------
    ws = wb.active; ws.title = "Summary"
    headline_row = 2 + int(list(wavg.index[wavg.WINDOW == HEADLINE])[0])
    tpos_row, ecl_row = 3 + len(wavg), 4 + len(wavg)
    rows = [
        ("As-of date", str(AS_OF), None),
        ("Cohorts (FY quarters)", len(cohorts), None),
        ("MOB grid", f"0..120 step 3  ({len(MOB_LIST)} points)", None),
        ("Loss-rate anchors", "84M and 120M", None),
        ("Headline window", HEADLINE, None),
        ("Weighted-avg loss rate", f"='Weighted_LR'!H{headline_row}", PC),
        ("Portfolio current TPOS (cr)", f"='Weighted_LR'!B{tpos_row}", CR),
        ("Portfolio ECL (cr) [provisional]", f"='Weighted_LR'!B{ecl_row}", CR),
        ("Final-ECL rule", "weighted-avg LR x current TPOS (multiplier unconfirmed)", None),
    ]
    for i, (k, v, fmt) in enumerate(rows):
        r = 1 + i
        kc = ws.cell(r, 1, k); kc.font, kc.fill, kc.border, kc.alignment = IFONT, IFL, BD, L
        vc = ws.cell(r, 2, v); vc.border, vc.alignment = BD, L
        if fmt: vc.number_format = fmt
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

    # 3-4) amount triangles (values) ------------------------------------------
    write_amount_triangle(wb.create_sheet("Pivot_90plus"), a90, disb)
    write_amount_triangle(wb.create_sheet("Pivot_TPOS"),   atp, disb)

    # 5) BadRate_90plus (PD formulas) -----------------------------------------
    write_badrate_triangle(wb.create_sheet("BadRate_90plus"), cohorts, "Pivot_90plus", "PD_")

    # 6) Movements (formula links) --------------------------------------------
    ws = wb.create_sheet("Movements")
    def mv_block(r0, prefix, src_sheet):
        for j, h in enumerate(["FY_QUARTER"] + [f"{prefix}{m}MOB" for m in ANCHORS], 1):
            c = ws.cell(r0, j, h); c.fill, c.font, c.alignment, c.border = HF, HFONT, C, BD
        for i, q in enumerate(cohorts):
            rr = r0 + 1 + i; src_r = row_of[q]
            ic = ws.cell(rr, 1, q); ic.fill, ic.font, ic.border = IFL, IFONT, BD
            for j, m in enumerate(ANCHORS):
                c = ws.cell(rr, 2 + j, f"='{src_sheet}'!{anchor_letter(m)}{src_r}")
                c.number_format, c.alignment, c.border = CR, C, BD
        return r0 + 1 + len(cohorts)
    end = mv_block(1, "TPOS_", "Pivot_TPOS")
    mv_block(end + 2, "90PLUS_", "Pivot_90plus")
    ws.column_dimensions["A"].width = 12

    # 7) LossRate_Qtr (formulas) ----------------------------------------------
    ws = wb.create_sheet("LossRate_Qtr")
    for j, h in enumerate(["FY_QUARTER", "DISB_AMT (weight)", "LOSS_RATE_84M", "LOSS_RATE_120M",
                           "CURRENT_MOB", "CURRENT_TPOS"], 1):
        c = ws.cell(1, j, h); c.fill, c.font, c.alignment, c.border = HF, HFONT, C, BD
    for i, rowq in qtr.iterrows():
        q = rowq.FY_QUARTER; rr = 2 + i; src_r = row_of[q]
        ic = ws.cell(rr, 1, q); ic.fill, ic.font = IFL, IFONT
        ws.cell(rr, 2, f"='Pivot_90plus'!$B{src_r}").number_format = CR
        for k, A in enumerate(ANCHOR_MOBS):
            num = f"'Pivot_90plus'!{anchor_letter(A)}{src_r}"
            den = "+".join(f"'Pivot_TPOS'!{anchor_letter(a)}{src_r}" for a in anchors_for(A))
            cc = ws.cell(rr, 3 + k, f"=IFERROR({num}/({den}),0)"); cc.number_format = PC
            if rowq[f"LOSS_RATE_{A}M"] > 1: cc.fill = WARN
        ws.cell(rr, 5, int(rowq.CURRENT_MOB))
        cm_letter = mob_letter(MOB_LIST.index(int(rowq.CURRENT_MOB)))
        ws.cell(rr, 6, f"='Pivot_TPOS'!{cm_letter}{src_r}").number_format = CR
        for cc in range(1, 7): ws.cell(rr, cc).alignment = C; ws.cell(rr, cc).border = BD
    last_lr_row = 1 + len(qtr)
    ws.column_dimensions["A"].width = 12
    for col in "BCDEF": ws.column_dimensions[col].width = 16

    # 8) Weighted_LR (formulas) -----------------------------------------------
    ws = wb.create_sheet("Weighted_LR")
    heads = ["WINDOW", "FY_START", "FY_END", "ANCHOR", "N_QTRS", "TOTAL_DISB", "SIMPLE_AVG", "WEIGHTED_AVG"]
    for j, h in enumerate(heads, 1):
        c = ws.cell(1, j, h); c.fill, c.font, c.alignment, c.border = HF, HFONT, C, BD
    lr_col = {84: "C", 120: "D"}
    for i, r in wavg.iterrows():
        rr = 2 + i
        k1, k2 = fy_key(r.FY_START), fy_key(r.FY_END)
        win_rows = [row_of[q] for q in cohorts if k1 <= fy_key(q) <= k2]
        r1, r2 = min(win_rows), max(win_rows)
        LRc = lr_col[int(r.ANCHOR_MOB)]
        disb_rng = f"LossRate_Qtr!$B{r1}:$B{r2}"
        lr_rng   = f"LossRate_Qtr!{LRc}{r1}:{LRc}{r2}"
        ws.cell(rr, 1, r.WINDOW); ws.cell(rr, 2, r.FY_START); ws.cell(rr, 3, r.FY_END)
        ws.cell(rr, 4, int(r.ANCHOR_MOB)); ws.cell(rr, 5, r2 - r1 + 1)
        ws.cell(rr, 6, f"=SUM({disb_rng})").number_format = CR
        ws.cell(rr, 7, f"=AVERAGE({lr_rng})").number_format = PC
        wc = ws.cell(rr, 8, f"=IFERROR(SUMPRODUCT({lr_rng},{disb_rng})/SUM({disb_rng}),0)")
        wc.number_format = PC; wc.font = Font(bold=True)
        if r.WEIGHTED_AVG_LR > 1: wc.fill = WARN
        for cc in range(1, 9): ws.cell(rr, cc).alignment = C; ws.cell(rr, cc).border = BD
    r0 = 3 + len(wavg)
    ws.cell(r0, 1, "Portfolio current TPOS (cr)").font = Font(bold=True)
    ws.cell(r0, 2, f"=SUM(LossRate_Qtr!F2:F{last_lr_row})").number_format = CR
    ws.cell(r0 + 1, 1, "Portfolio ECL (cr) [provisional]").font = Font(bold=True)
    ec = ws.cell(r0 + 1, 2, f"=H{headline_row}*B{r0}"); ec.number_format = CR; ec.fill = TOT
    for col, w in zip("ABCDEFGH", [20, 10, 10, 9, 9, 14, 13, 14]):
        ws.column_dimensions[col].width = w

    wb.save(path)
    return wb


if __name__ == "__main__":
    from types import SimpleNamespace
    feed = pd.read_csv("DATA_ECL_NEW.csv")
    a90 = pd.read_csv("tri_90plus_amt.csv", index_col=0);  a90.columns = [int(c) for c in a90.columns]
    atp = pd.read_csv("tri_tpos_amt.csv",  index_col=0);  atp.columns = [int(c) for c in atp.columns]
    disb = feed.groupby("FY_QUARTER").DISBURSAL_AMT.sum().reindex(a90.index)
    qtr  = pd.read_csv("ecl_by_quarter.csv")
    wavg = pd.read_csv("weighted_loss_rate.csv")
    tris = SimpleNamespace(a90=a90, atp=atp, disb=disb)
    ecl  = SimpleNamespace(by_quarter=qtr, wavg=wavg)
    wb = build_excel(feed, tris, SimpleNamespace(), ecl, OUT)
    print("REPORT COMPLETE ->", OUT, "| sheets:", wb.sheetnames)