"""
config.py - SINGLE SOURCE OF TRUTH for every shared knob.

Each pipeline script begins with `from config import *` and defines no shared
constants of its own. To roll the quarter you change AS_OF here and nothing else.

run_all.py does NOT import this - it only shells out to the scripts.
Per-file Excel styles (HF, BD, C, ...) stay in their own scripts; they are
presentation, not configuration.
"""
from datetime import date

# THE ONE KNOB.  Reporting as-of ("today"). Drives triangle maturity, which
# cells get projected, and the current-TPOS diagonal. Roll this each quarter.
AS_OF = date(2026, 6, 30)

# disbursal window (last ~10 years).
# END_DISB is DERIVED from AS_OF so the two can never silently disagree.
START      = date(2015, 4, 1)
START_DISB = START.isoformat()
END_DISB   = AS_OF.isoformat()

# MOB grid: 0,3,6,...,120  (41 points)
MOB_LIST = list(range(0, 121, 3))

# loss-rate anchors.  LR_A = 90+ @ A / SUM(TPOS 12,24,...,A)
ANCHOR_MOBS = [84, 120]
anchors_for = lambda a: list(range(12, a + 1, 12))
ALL_ANCHORS = anchors_for(max(ANCHOR_MOBS))     # 12..120, used by movement tables
ANCHORS     = ALL_ANCHORS                       # alias used by report.py

# observation windows for the disbursal-weighted average loss rate:
#   (label, fy_start, fy_end, anchor)
WINDOWS = [
    ("FY20-FY23 @ 84M",  "FY20-Q1", "FY23-Q4",  84),
    ("FY16-FY23 @ 84M",  "FY16-Q1", "FY23-Q4",  84),
    ("FY16-FY23 @ 120M", "FY16-Q1", "FY23-Q4", 120),
]
HEADLINE = "FY20-FY23 @ 84M"    # window used for the provisional ECL number

# None = whole book (all segments); or 1..5 for a single segment
SEGMENT = None

# synthetic data generation (base_loans.py only)
SEED     = 34
N_LOANS  = 60_000
SEGMENTS = [1, 2, 3, 4, 5]
SEG_DEFAULT_RATE = {1: 0.020, 2: 0.035, 3: 0.055, 4: 0.080, 5: 0.120}

# If True, a defaulted loan's TPOS freezes at outstanding-at-default instead of
# continuing to amortise. More realistic, but breaks exact reconciliation with
# the original wide CSV. Keep False unless the mentor asks otherwise.
FREEZE_TPOS_AT_DEFAULT = False

# validation
TOL = 1e-4                      # crore tolerance

# file names
DB_PATH   = "ecl.db"
BASE_CSV  = "base_loans.csv"
PERF_CSV  = "performance_long.csv"
FEED_CSV  = "data_ecl.csv"
OUT_CSV   = FEED_CSV            # alias used by sql_refactor.py
OUT_SQL   = "phase1_generated.sql"
TRI_90    = "tri_90plus_amt.csv"
TRI_TP    = "tri_tpos_amt.csv"
TRI_90_R  = "tri_90plus_rate.csv"
TRI_TP_R  = "tri_tpos_rate.csv"
MV_90     = "movement_90plus.csv"
MV_TP     = "movement_tpos.csv"
LOSS_CSV  = "loss_rate.csv"
WAVG_CSV  = "weighted_loss_rate.csv"
QTR_CSV   = "ecl_by_quarter.csv"
REPORT_XLSX     = "ECL_Report.xlsx"
VALIDATION_XLSX = "validation_report.xlsx"
OUT       = REPORT_XLSX         # alias used by report.py

# shared helper: FY-quarter sort key   'FY16-Q1' -> (16, 1)
fy_key = lambda q: (int(q[2:4]), int(q[-1]))