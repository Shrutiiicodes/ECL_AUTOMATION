"""
MOVEMENT TABLES + LOSS RATES  (84M and 120M anchors)
====================================================
Consumes the completed amount triangles and produces, per FY_QUARTER:

  movement_tpos    : TPOS at the anchor MOBs
  movement_90plus  : 90+ settlement at the anchor MOBs
  loss table       : loss rate at EACH anchor

Loss rate at anchor A:
    LR_A(q) = 90+(q, A) / SUM( TPOS(q, 12), TPOS(q, 24), ..., TPOS(q, A) )

  84M  -> anchors 12,24,36,48,60,72,84          (7 terms)   == the bank's =I97/SUM(B51:H51)
  120M -> anchors 12,24,...,108,120             (10 terms)

Anchors are MOB LEVELS at those ages (not deltas), matching Movements.
All amounts in crores. DISBURSAL_AMT is carried through because it is the WEIGHT
used for the weighted-average loss rate in the next stage.

-------------------------------------------------------------------------------
This module exposes a PURE function:

    run(tri_90_amt, tri_tpos_amt, feed) -> LossRates(loss, mv90, mvtp)

It reads nothing from disk and writes nothing to disk. The orchestrator passes
DataFrames in and gets DataFrames out. CSV/Excel side effects live only in the
`if __name__ == "__main__"` block below, which exists purely so the phase can
still be run and inspected standalone.
"""

from typing import NamedTuple

import numpy as np
import pandas as pd

from src.config import *      # ALL_ANCHORS, ANCHOR_MOBS, anchors_for, and (for __main__) TRI_90, TRI_TP, FEED_CSV


class LossRates(NamedTuple):
    loss: pd.DataFrame   # FY_QUARTER, DISBURSAL_AMT, and per-anchor NINETY_PLUS_/TPOS_SUM_12_/LOSS_RATE_
    mv90: pd.DataFrame   # 90+ movement at the anchor MOBs
    mvtp: pd.DataFrame   # TPOS movement at the anchor MOBs


def run(tri_90_amt: pd.DataFrame, tri_tpos_amt: pd.DataFrame, feed: pd.DataFrame) -> LossRates:
    """Compute movement tables and per-anchor loss rates. No I/O."""
    # Triangle columns may arrive as ints (in-memory from chain_ladder) or as
    # strings (round-tripped through CSV in standalone mode). Normalise on copies
    # so we never mutate the caller's frames.
    a90 = tri_90_amt.copy(); a90.columns = [int(c) for c in a90.columns]
    atp = tri_tpos_amt.copy(); atp.columns = [int(c) for c in atp.columns]

    disb = feed.groupby("FY_QUARTER").DISBURSAL_AMT.sum().reindex(a90.index)

    mv90 = a90[ALL_ANCHORS].copy()
    mvtp = atp[ALL_ANCHORS].copy()

    # LOSS RATE PER ANCHOR   (IFERROR -> 0)
    loss = pd.DataFrame({"FY_QUARTER": a90.index, "DISBURSAL_AMT": disb.values})
    for A in ANCHOR_MOBS:
        ancs = anchors_for(A)
        num = a90[A]
        den = atp[ancs].sum(axis=1)
        lr = (num / den).replace([np.inf, -np.inf], 0).fillna(0)
        lr = lr.where(den != 0, 0.0)
        loss[f"NINETY_PLUS_{A}"] = num.values
        loss[f"TPOS_SUM_12_{A}"] = den.values
        loss[f"LOSS_RATE_{A}M"]  = lr.values

    return LossRates(loss=loss, mv90=mv90, mvtp=mvtp)





def _print_summary(res: LossRates, tri_90_amt: pd.DataFrame) -> None:
    loss, mvtp = res.loss, res.mvtp
    print("=" * 60); print("LOSS RATES COMPLETE"); print("=" * 60)
    print(f"quarters              : {len(loss)}")
    for A in ANCHOR_MOBS:
        col = loss[f"LOSS_RATE_{A}M"]
        print(f"  {A:>3}M anchors {str(anchors_for(A)):<44}")
        print(f"       min/median/max = {col.min():.4%} / {col.median():.4%} / {col.max():.4%}"
              f"   |  >100%: {int((col > 1).sum())}")

    q = "FY18-Q1"
    if q in loss.FY_QUARTER.values:
        r = loss[loss.FY_QUARTER == q].iloc[0]
        hd = float(mvtp.loc[q, anchors_for(84)].sum()); hn = float(tri_90_amt.loc[q, 84])
        print(f"\nhand-check {q} @84M: {hn:.6f} / {hd:.6f} = {hn/hd:.6%}  (engine {r.LOSS_RATE_84M:.6%})")


if __name__ == "__main__":
    a90 = pd.read_csv(TRI_90, index_col=0); a90.columns = [int(c) for c in a90.columns]
    atp = pd.read_csv(TRI_TP, index_col=0); atp.columns = [int(c) for c in atp.columns]
    feed = pd.read_csv(FEED_CSV)

    res = run(a90, atp, feed)

    res.mv90.to_csv(MV_90)
    res.mvtp.to_csv(MV_TP)
    res.loss.to_csv(LOSS_CSV, index=False)
    _print_summary(res, a90)
    print(f"\nWrote: {MV_90}, {MV_TP}, {LOSS_CSV}")