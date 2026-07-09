"""
config.py - single source of truth for the shared knobs.
Change the quarter here instead of editing each script.

Recommended refactor: replace the inline CONFIG block at the top of each script
with `from config import *`. The scripts currently carry these values inline for
standalone clarity, so a run works as-is; centralising here makes the quarterly
change a one-file edit.
"""
from datetime import date

# reporting as-of ("today"): drives triangle maturity + current-TPOS diagonal
AS_OF = date(2026, 7, 7)

# disbursal window to include (last ~10 years)
START_DISB = "2015-04-01"
END_DISB   = "2026-07-07"

# MOB grid: 0,3,6,...,120  (41 points)
MOB_LIST = list(range(0, 121, 3))

# loss-rate anchors. LR_A = 90+ @ A / SUM(TPOS 12,24,...,A)
ANCHOR_MOBS = [84, 120]
anchors_for = lambda a: list(range(12, a + 1, 12))

# observation windows for the disbursal-weighted average loss rate:
#   (label, fy_start, fy_end, anchor)
WINDOWS = [
    ("FY20-FY23 @ 84M",  "FY20-Q1", "FY23-Q4",  84),
    ("FY16-FY23 @ 84M",  "FY16-Q1", "FY23-Q4",  84),
    ("FY16-FY23 @ 120M", "FY16-Q1", "FY23-Q4", 120),
]
HEADLINE = "FY20-FY23 @ 84M"      # window used for the provisional ECL number

# None = whole book (all segments); or 1..5 to run a single segment
SEGMENT = None

# data generation
SEED    = 34
N_LOANS = 60_000

# validation tolerance (crores)
TOL = 1e-4

DB_PATH = "ecl.db"