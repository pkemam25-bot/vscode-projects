import datetime
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

# Month boundaries as Excel serials (Jan 2026 = 46023)
def serial(y, m, d):
    return (datetime.date(y, m, d) - datetime.date(1899, 12, 30)).days

MONTHS = [
    ("Jan26", serial(2026, 1, 1), serial(2026, 2, 1)),
    ("Feb26", serial(2026, 2, 1), serial(2026, 3, 1)),
    ("Mar26", serial(2026, 3, 1), serial(2026, 4, 1)),
    ("Apr26", serial(2026, 4, 1), serial(2026, 5, 1)),
    ("May26", serial(2026, 5, 1), serial(2026, 6, 1)),
    ("Jun26", serial(2026, 6, 1), serial(2026, 7, 1)),
    ("Jul26", serial(2026, 7, 1), serial(2026, 8, 1)),
    ("Aug26", serial(2026, 8, 1), serial(2026, 9, 1)),
    ("Sep26", serial(2026, 9, 1), serial(2026, 10, 1)),
    ("Oct26", serial(2026, 10, 1), serial(2026, 11, 1)),
    ("Nov26", serial(2026, 11, 1), serial(2026, 12, 1)),
    ("Dec26", serial(2026, 12, 1), serial(2027, 1, 1)),
    ("Jan27", serial(2027, 1, 1), serial(2027, 2, 1)),
]

# (label, start_serial, end_serial, milestone?)
ACTIVITIES = [
    ("Consolidate Requirement",            serial(2026,2,16), serial(2026,3,13), False),
    ("Design architecture",                serial(2026,3,16), serial(2026,4,27), False),
    ("SOLAR + ARB",                        serial(2026,3,30), serial(2026,4,10), False),
    ("Prepare data, sharing & permission", serial(2026,3,2),  serial(2026,5,29), False),
    ("Model Dev & Prediction Pipeline",    serial(2026,3,2),  serial(2026,5,29), False),
    ("Deploy Group 1 (1-25 wells)",        serial(2026,6,1),  serial(2026,7,3),  True),
    ("Feedback/improvement Group 1",       serial(2026,7,6),  serial(2026,7,31), False),
    ("Model Adjustment",                   serial(2026,7,27), serial(2026,10,2), False),
    ("Deploy Group 2 (26-120 wells)",      serial(2026,10,5), serial(2026,10,30), True),
    ("Feedback/improvement Group 2",       serial(2026,11,2), serial(2026,11,20), False),
    ("Deploy Rest (121-160 wells)",        serial(2026,11,23), serial(2026,12,18), True),
    ("Feedback/improvement Rest",          serial(2026,12,21), serial(2027,1,1),  False),
    ("IPOS Phase 2: Integration Support",  serial(2026,6,1),  serial(2026,8,14),  False),
    ("Handover: manuals & training",       serial(2026,12,21), serial(2027,1,1),  True),
    ("Operation Support (Tier 1,2)",       serial(2026,1,4),  serial(2026,4,10),  False),
]

W = len(MONTHS)

def render():
    # header
    hdr = " " * 34 + "".join(m[0].ljust(9) for m in MONTHS)
    print(hdr)
    print(" " * 34 + "-" * (W * 9))
    for label, start, end, ms in ACTIVITIES:
        line = label.ljust(33) + " "
        for i, (name, s, e) in enumerate(MONTHS):
            cell = ""
            # overlap of [start,end] with [s,e)
            if end > s and start < e:
                cell = "█" * 9
            if ms and start >= s and start < e:
                cell = cell[:4] + "♦" + cell[5:]
            line += cell.ljust(9)
        print(line)
    print()
    print("♦ = milestone start")

render()
