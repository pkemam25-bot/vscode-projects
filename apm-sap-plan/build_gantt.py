"""Build 'APM SAP Plan - Gantt.xlsx' from 'APM SAP Plan.xlsx'.

TWO-TAB workbook (two stakeholder groups):

Tab 'SAP Timeline'  - for SAP-only stakeholders. One row per activity,
    phase-colored SAP bars ONLY (no APM rows, no FTE). Status + SAP dates are
    LINKED formulas to the master tab, so this view is read-only and always
    in sync. Week columns stay narrow (width 2).

Tab 'APM Plan & FTE' - the editable master (APM team). Each activity has TWO
    stacked band rows: phase color = SAP window (SAP Start/End D:E) and
    blue = APM's own window (APM Start/End F:G; red TBC = not confirmed yet).
    Under the Gantt, an FTE tracking grid: one block per activity with role
    rows x week columns (MOCK sample values inside the APM window, editable),
    weekly subtotal formulas, and light-blue shading that follows the APM
    window (conditional formatting). Week columns wider (~5) so numbers show.

Shared behaviour (both tabs):
  * Week grid (Mondays), month bands over week start days (row 4/5 headers).
  * Bars are CONDITIONAL FORMATTING keyed to the date columns -> live repaint.
  * Today marker (current week): orange header cell + light gray band.
  * Go-Live milestone row (editable date on master; red block follows).
  * Status dropdown + date pickers on the master tab.
  * Previous generated file is backed up (timestamped) before every run.
The original source file is never modified.
"""

import calendar
import re
import shutil
from datetime import date, datetime, timedelta
from pathlib import Path

import openpyxl
from openpyxl.formatting.rule import Rule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

BASE = Path(__file__).resolve().parent
SRC = BASE / "APM SAP Plan.xlsx"
OUT = BASE / "APM SAP Plan - Gantt.xlsx"
SHEET = "SAP APM Plan"

MONTHS = {
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
    "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12,
}
SHORT_MONTH = ["", "Jan", "Feb", "Mar", "Apr", "May", "Jun",
               "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

# Source rows -> section letters / activities (row in source sheet, WBS number)
SECTION_ROWS = {3: "A", 5: "B", 10: "C", 14: "D", 18: "E"}
ACTIVITY_ROWS = [
    (4, 1), (6, 2), (7, 3), (8, 4), (9, 5),
    (11, 6), (12, 7), (13, 8),
    (15, 9), (16, 10), (17, 11),
    (19, 12),
]

# ---- colors ----
PHASE = {
    "A": dict(solid="9DC3E6", border="2F5597"),
    "B": dict(solid="A9D08E", border="538135"),
    "C": dict(solid="FFD966", border="BF9000"),
    "D": dict(solid="F4B183", border="C55A11"),
    "E": dict(solid="B4A7D6", border="5B4D92"),
}
APM_FILL = "4472C4"          # APM's own-timeline bar (distinct from phase colors)
APM_BORDER = "1F3864"
SECTION_FILL = "D9D9D9"
HEADER_FILL = "F2F2F2"
MILESTONE_FILL = "C00000"
TODAY_TINT = "E7E6E6"           # light gray band down the current week
WEEKLY_TODAY_FILL = "FFC000"    # orange current-week header
WEEKLY_TODAY_FONT = "1F1F1F"
TITLE_FILL = "1F3864"           # dark blue title bar
SUBTITLE_FILL = "EFEFEF"
FTE_HEADER_FILL = "BDD7EE"      # FTE block header band
FTE_SHADE = "DEEAF6"            # light blue = APM window on FTE grid
FTE_SUBTOTAL_FILL = "F2F2F2"

GRID_START = (2026, 3)          # first month covered (margin before Apr 2026)
GRID_END = (2027, 12)           # last month covered (after hypercare Sep 2027)
GO_LIVE = date(2027, 7, 1)

# Position rows (one row per person) with monthly FTE sample values Apr-Dec 26,
# taken from the user's FTE example. 2027 Jan-Sep mirrors the same pattern;
# Feb/Mar 26 and Oct-Dec 27 are 0.0 (program ramp-up / ramp-down margins).
POSITIONS = [
    ("ITPM",                  [0.20, 0.10, 0.11, 0.11, 0.10, 0.20, 0.20, 0.20, 0.20]),
    ("Scrum Master",          [0.40, 0.40, 0.40, 0.40, 0.40, 0.40, 0.40, 0.40, 0.40]),
    ("BA",                    [0.95, 0.95, 0.50, 0.50, 0.95, 1.00, 1.00, 1.00, 1.00]),
    ("UX/UI Designer",        [1.00, 1.00, 1.00, 1.00, 1.00, 1.00, 1.00, 1.00, 1.00]),
    ("Frontend Developer",    [0.85, 0.90, 0.90, 0.85, 0.90, 1.00, 1.00, 1.00, 1.00]),
    ("Frontend Developer",    [1.00, 1.00, 1.00, 1.00, 1.00, 1.00, 1.00, 1.00, 1.00]),
    ("Frontend Developer",    [1.00, 1.00, 1.00, 1.00, 1.00, 1.00, 1.00, 1.00, 1.00]),
    ("Backend Developer",     [1.00, 1.00, 1.00, 1.00, 1.00, 1.00, 1.00, 1.00, 1.00]),
    ("Backend Developer",     [0.75, 0.90, 0.90, 0.75, 0.90, 0.65, 0.00, 0.00, 0.00]),
    ("QA",                    [0.75, 0.90, 0.90, 0.75, 0.90, 0.80, 1.00, 1.00, 1.00]),
    ("QA",                    [1.00, 1.00, 1.00, 1.00, 1.00, 0.00, 0.00, 0.00, 0.00]),
    ("QA Automation",         [0.00, 0.00, 0.00, 0.00, 0.00, 0.00, 0.00, 0.00, 0.00]),
    ("Data Analytics Services", [0.20, 0.20, 0.20, 0.20, 0.10, 0.10, 0.00, 0.00, 0.00]),
    ("Data Scientist",        [0.00, 0.00, 0.00, 0.00, 0.00, 0.00, 0.00, 0.00, 0.00]),
]

MASTER_WEEK_WIDTH = 5.5        # APM Plan & FTE tab: numbers (FTE) must display
VIEW_WEEK_WIDTH = 2            # SAP Timeline tab: bars only, user preference

DATA_START_ROW = 6              # first data row (after headers)


def clean(text):
    """Source has non-breaking spaces and some mangled en dashes (U+FFFD)."""
    if text is None:
        return ""
    return str(text).replace("\xa0", " ").replace("\ufffd", "-")


def last_day(y, m):
    return calendar.monthrange(y, m)[1]


def parse_dates(text):
    """Parse '23-28 Sep 2026' style text into (start, end) or None (TBC/-/n/a)."""
    s = clean(text).strip()
    if not s or s.lower() in ("tbc", "-", "n/a") or "tbc" in s.lower():
        return None
    dsh = r"(?:-|\u2013|\u2014)"
    d = r"(\d{1,2})"
    mon = r"([A-Za-z]{3,9})"
    y4 = r"(\d{4})"
    y2 = r"(\d{2})"

    # 1) day-day same month-year:  '23-28 Sep 2026'   '26 - 30 Jun 2027'
    m = re.match(rf"^{d}\s*{dsh}\s*{d}\s+{mon}\s+{y4}$", s)
    if m:
        d1, d2, mo, yr = int(m.group(1)), int(m.group(2)), MONTHS[m.group(3).lower()[:3]], int(m.group(4))
        return date(yr, mo, d1), date(yr, mo, d2)

    # 2) explicit full dates, year on one or both sides:
    #    '1 Oct - 20 Dec 2026'   '3 May 2027 - 25 Jun 2027'   '4 Jan - 15 Mar 2027'
    m = re.match(rf"^{d}\s+{mon}(?:\s+{y4})?\s*{dsh}\s*{d}\s+{mon}\s+{y4}$", s)
    if m:
        sd, sm = int(m.group(1)), MONTHS[m.group(2).lower()[:3]]
        sy = int(m.group(3)) if m.group(3) else None
        ed, em = int(m.group(4)), MONTHS[m.group(5).lower()[:3]]
        ey = int(m.group(6))
        sy = sy or ey
        return date(sy, sm, sd), date(ey, em, ed)

    # 3) month year - month year:  'Dec 2026 - Mar 2027'
    m = re.match(rf"^{mon}\s+{y4}\s*{dsh}\s*{mon}\s+{y4}$", s)
    if m:
        y1, m1 = int(m.group(2)), MONTHS[m.group(1).lower()[:3]]
        y2, m2 = int(m.group(4)), MONTHS[m.group(3).lower()[:3]]
        return date(y1, m1, 1), date(y2, m2, last_day(y2, m2))

    # 4) 'Apr 26' style 2-digit year:  'Apr 26 - Sep 2027'
    m = re.match(rf"^{mon}\s+{y2}\s*{dsh}\s*{mon}\s+{y4}$", s)
    if m:
        y1, m1 = 2000 + int(m.group(2)), MONTHS[m.group(1).lower()[:3]]
        y2, m2 = int(m.group(4)), MONTHS[m.group(3).lower()[:3]]
        return date(y1, m1, 1), date(y2, m2, last_day(y2, m2))

    # 5) month - month year:  'Apr - Jun 2026'   'Apr - Nov 2026'
    m = re.match(rf"^{mon}\s*{dsh}\s*{mon}\s+{y4}$", s)
    if m:
        y, m1 = int(m.group(3)), MONTHS[m.group(1).lower()[:3]]
        m2 = MONTHS[m.group(2).lower()[:3]]
        return date(y, m1, 1), date(y, m2, last_day(y, m2))

    # 6) single date (APM col F uses e.g. '9-Oct-26'): start == end
    m = re.match(rf"^{d}\s*{dsh}\s*{mon}\s*{dsh}\s*({y2}|{y4})$", s)
    if m:
        dd, mo = int(m.group(1)), MONTHS[m.group(2).lower()[:3]]
        yr = int(m.group(3))
        yr = yr if yr > 100 else 2000 + yr
        return date(yr, mo, dd), date(yr, mo, dd)

    return None


def load_rows():
    """Read the source sheet into ordered row dicts (section headers + activities)."""
    wb = openpyxl.load_workbook(SRC, data_only=True)
    ws = wb[SHEET]
    items = []
    for src_row, letter in SECTION_ROWS.items():
        items.append((src_row, {"kind": "section", "letter": letter,
                                "label": clean(ws.cell(row=src_row, column=1).value)}))
    for src_row, num in ACTIVITY_ROWS:
        name = clean(ws.cell(row=src_row, column=2).value)
        status = clean(ws.cell(row=src_row, column=3).value)
        timeline = clean(ws.cell(row=src_row, column=4).value)          # SAP window
        parsed = parse_dates(timeline)
        # col F 'Involved APM Plan (approx.)': APM's own window (many TBC).
        # The cell may be a real Excel date (Mock 2 = 9-Oct-26) or free text.
        raw_apm = ws.cell(row=src_row, column=6).value
        apm_raw = clean(raw_apm).strip()
        if isinstance(raw_apm, datetime):
            apm_parsed = (raw_apm.date(), raw_apm.date())
        else:
            apm_parsed = parse_dates(apm_raw)
        # #1 'Risk, action, issue tracking' monitors the whole lifecycle -> mirror SAP
        if num == 1 and apm_parsed is None and parsed:
            apm_parsed = parsed
        # col E 'Mock Readiness': dates for Mock 1-3, 'TBC', or not required
        raw_mock = ws.cell(row=src_row, column=5).value
        cleaned_mock = clean(raw_mock).strip()
        if isinstance(raw_mock, datetime):
            mock = ("date", raw_mock.date())
        elif cleaned_mock.lower() in ("tbc", "-"):
            mock = ("text", "TBC")
        else:
            mock = None
        # col H 'Remark' (free text)
        remark = clean(ws.cell(row=src_row, column=8).value)
        items.append((src_row, {
            "kind": "activity", "num": num, "name": name, "status": status,
            "timeline_raw": timeline, "apm_raw": apm_raw, "mock": mock,
            "remark": remark,
            "start": parsed[0] if parsed else None,
            "end": parsed[1] if parsed else None,
            "apm_start": apm_parsed[0] if apm_parsed else None,
            "apm_end": apm_parsed[1] if apm_parsed else None,
        }))
    rows = [entry for _, entry in sorted(items, key=lambda t: t[0])]
    cur = None
    for row in rows:
        if row["kind"] == "section":
            cur = row["letter"]
        else:
            row["phase"] = cur
    return rows


def solid_fill(hexcolor):
    return PatternFill(start_color=hexcolor, end_color=hexcolor, fill_type="solid")


def calc_weeks():
    """Monday-start weeks covering the whole grid (for both tabs)."""
    first = date(GRID_START[0], GRID_START[1], 1)
    first_monday = first - timedelta(days=first.weekday())
    weeks = []
    w = first_monday
    while w <= date(GRID_END[0], GRID_END[1], 31):
        weeks.append(w)
        w += timedelta(days=7)
    return weeks


def month_groups(weeks):
    """(start_idx, end_idx, year, month) runs over the week list."""
    groups = []
    start, key = 0, (weeks[0].year, weeks[0].month)
    for i in range(1, len(weeks) + 1):
        if i == len(weeks) or (weeks[i].year, weeks[i].month) != key:
            groups.append((start, i - 1, key[0], key[1]))
            if i < len(weeks):
                start, key = i, (weeks[i].year, weeks[i].month)
    return groups


def backup_previous():
    if OUT.exists():
        ts = datetime.now().strftime("%Y%m%d-%H%M%S")
        bak = OUT.with_name(f"{OUT.stem} backup {ts}{OUT.suffix}")
        shutil.copy2(OUT, bak)
        print(f"[backup] previous version -> {bak.name}")


def write_mock(ws, r, row, col):
    """'Mock Readiness': readiness dates for Mock 1-3, grey hatch elsewhere."""
    cell = ws.cell(row=r, column=col)
    cell.alignment = Alignment(horizontal="center", vertical="center")
    mock = row.get("mock")
    if mock is None:
        cell.fill = PatternFill(patternType="lightUp", fgColor="BFBFBF", bgColor="FFFFFF")
    elif mock[0] == "date":
        cell.value = mock[1]
        cell.number_format = "mmm yy" if mock[1].day == 1 else "dd-mmm-yy"
        cell.font = Font(bold=True, size=9)
    else:
        cell.value = mock[1]
        cell.font = Font(italic=True, bold=True, color="C00000", size=9)


def week_headers(ws, weeks, mfc, last_col_letter):
    """Month bands (row 4, merged) over week start days (row 5)."""
    thin_gray = Side(style="thin", color="BFBFBF")
    for i, wd in enumerate(weeks):
        col = mfc + i
        c = ws.cell(row=5, column=col, value=wd)     # real date underneath...
        c.number_format = "d"                        # ...displayed as day-of-month only
        c.font = Font(bold=True, size=8)
        c.alignment = Alignment(horizontal="center")
        c.fill = solid_fill(HEADER_FILL)
        c.border = Border(left=thin_gray, right=thin_gray, top=thin_gray, bottom=thin_gray)
    for g_start, g_end, gy, gm in month_groups(weeks):
        c0, c1 = mfc + g_start, mfc + g_end
        if c1 > c0:
            ws.merge_cells(start_row=4, start_column=c0, end_row=4, end_column=c1)
        label = SHORT_MONTH[gm] + (f" {str(gy)[2:]}" if gm == 1 else "")
        c = ws.cell(row=4, column=c0, value=label)
        c.font = Font(bold=True, size=9)
        c.alignment = Alignment(horizontal="center", vertical="center")
        for col in range(c0, c1 + 1):
            ws.cell(row=4, column=col).fill = solid_fill(HEADER_FILL)
    return thin_gray


def fixed_headers(ws, headers, thin_gray):
    for i, h in enumerate(headers, start=1):
        c = ws.cell(row=5, column=i, value=h)
        c.font = Font(bold=True, size=9)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.fill = solid_fill(HEADER_FILL)
        c.border = Border(left=thin_gray, right=thin_gray, top=thin_gray, bottom=thin_gray)


def titles(ws, last_col, title, subtitle):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=last_col)
    t = ws.cell(row=1, column=1, value=title)
    t.fill = solid_fill(TITLE_FILL)
    t.font = Font(bold=True, size=15, color="FFFFFF")
    t.alignment = Alignment(vertical="center")
    ws.row_dimensions[1].height = 26

    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=last_col)
    sub = ws.cell(row=2, column=1, value=subtitle)
    sub.fill = solid_fill(SUBTITLE_FILL)
    sub.font = Font(italic=True, size=9, color="404040")
    sub.alignment = Alignment(vertical="center")
    ws.row_dimensions[2].height = 34


def bar_cf(ws, row, mfc, last_col, dref, eref, fill_hex, border_hex):
    """Expression CF: bar when [dref, eref] overlaps the (relative) week column."""
    c = get_column_letter(mfc)
    end = get_column_letter(last_col)
    side = Side(style="thin", color=border_hex)
    dxf = DifferentialStyle(
        fill=PatternFill(start_color=fill_hex, end_color=fill_hex, fill_type="solid"),
        border=Border(left=side, right=side, top=side, bottom=side))
    rule = Rule(type="expression", dxf=dxf, stopIfTrue=True)
    rule.formula = [f"AND(${dref}<={c}$5+6,${eref}>={c}$5)"]
    ws.conditional_formatting.add(f"{c}{row}:{end}{row}", rule)


def tint_cf(ws, first_row, last_row, mfc, last_col):
    c = get_column_letter(mfc)
    end = get_column_letter(last_col)
    dxf = DifferentialStyle(fill=PatternFill(start_color=TODAY_TINT, end_color=TODAY_TINT,
                                             fill_type="solid"))
    rule = Rule(type="expression", dxf=dxf)
    rule.formula = [f"AND({c}$5<=TODAY(),{c}$5+6>=TODAY())"]
    ws.conditional_formatting.add(f"{c}{first_row}:{end}{last_row}", rule)


def today_header_cf(ws, mfc, last_col):
    c = get_column_letter(mfc)
    end = get_column_letter(last_col)
    dxf = DifferentialStyle(
        fill=PatternFill(start_color=WEEKLY_TODAY_FILL, end_color=WEEKLY_TODAY_FILL,
                         fill_type="solid"),
        font=Font(bold=True, color=WEEKLY_TODAY_FONT, size=8))
    rule = Rule(type="expression", dxf=dxf)
    rule.formula = [f"AND({c}5<=TODAY(),{c}5+6>=TODAY())"]
    ws.conditional_formatting.add(f"{c}5:{end}5", rule)


def milestone_row(ws, row, mfc, last_col, date_value):
    """Dark milestone row; red block follows the date in col D of this row."""
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)
    mc = ws.cell(row=row, column=1, value="SAP Go-Live (Milestone)")
    mc.font = Font(bold=True, size=10, color="FFFFFF")
    mc.alignment = Alignment(vertical="center")
    for col in range(1, last_col + 1):
        ws.cell(row=row, column=col).fill = solid_fill("404040")
    gc = ws.cell(row=row, column=4, value=date_value)
    gc.number_format = "dd mmm yy"
    gc.alignment = Alignment(horizontal="center")
    gc.font = Font(color="FFFFFF")
    c = get_column_letter(mfc)
    end = get_column_letter(last_col)
    red_side = Side(style="thin", color="7F0000")
    dxf = DifferentialStyle(
        fill=PatternFill(start_color=MILESTONE_FILL, end_color=MILESTONE_FILL,
                         fill_type="solid"),
        border=Border(left=red_side, right=red_side, top=red_side, bottom=red_side))
    rule = Rule(type="expression", dxf=dxf)
    rule.formula = [f"AND({c}$5<=$D{row},{c}$5+6>=$D{row})"]
    ws.conditional_formatting.add(f"{c}{row}:{end}{row}", rule)
    ws.row_dimensions[row].height = 18
    return row


def window_text(row):
    if row["apm_start"] and row["apm_end"]:
        return (f"{row['apm_start']:%d %b %y} - {row['apm_end']:%d %b %y}")
    return "TBC - pending confirmation"


# =====================================================================
#  TAB 2 (master): 'APM Plan & FTE'
# =====================================================================
def build_master(wb, rows, weeks):
    ws = wb.active
    ws.title = "APM Plan & FTE"
    ws.sheet_view.showGridLines = False

    n = len(weeks)
    mfc = 11                       # fixed block A..J, weeks start col K
    first_letter = get_column_letter(mfc)
    last_col = (mfc - 1) + n
    last_letter = get_column_letter(last_col)

    for c, wd in {1: 4.5, 2: 62, 3: 12, 4: 13, 5: 13, 6: 13, 7: 13,
                  8: 10, 9: 13, 10: 45}.items():
        ws.column_dimensions[get_column_letter(c)].width = wd
    for i in range(n):
        ws.column_dimensions[get_column_letter(mfc + i)].width = MASTER_WEEK_WIDTH

    titles(ws, last_col, "APM - SAP Activity & Support Plan  |  Gantt (Weekly) + FTE Tracking",
           "Two stacked bars per activity: phase color = SAP window (SAP Start/End D:E); "
           "blue = APM's own window (APM Start/End F:G; red TBC = not confirmed yet) | "
           "Edit dates/Status/FTE here - bars, Dur and FTE shading update automatically | "
           "FTE grid below: role rows x weeks, values inside each activity's APM window | "
           "orange header + gray band = current week; red block = Go-Live week | "
           "AS OF 27 AUG 2026 - SAP Go-Live 1 Jul 2027")

    # legend (row 3)
    leg = ws.cell(row=3, column=1, value="Phase:")
    leg.font = Font(bold=True, size=9)
    leg.alignment = Alignment(horizontal="right", vertical="center")
    thin_dark = Side(style="thin", color="404040")
    for i, letter in enumerate(PHASE):
        chip = ws.cell(row=3, column=2 + i, value=letter)
        chip.fill = solid_fill(PHASE[letter]["solid"])
        chip.font = Font(bold=True, size=9, color="1F1F1F")
        chip.alignment = Alignment(horizontal="center", vertical="center")
        chip.border = Border(left=thin_dark, right=thin_dark, top=thin_dark, bottom=thin_dark)
    apm_chip = ws.cell(row=3, column=8, value="")
    apm_chip.fill = solid_fill(APM_FILL)
    apm_chip.border = Border(left=thin_dark, right=thin_dark, top=thin_dark, bottom=thin_dark)
    ws.merge_cells(start_row=3, start_column=9, end_row=3, end_column=10)
    hl = ws.cell(row=3, column=9, value="APM own window (cols F-G)")
    hl.font = Font(italic=True, size=8, color="595959")
    hl.alignment = Alignment(vertical="center")
    ws.row_dimensions[3].height = 14

    thin_gray = week_headers(ws, weeks, mfc, last_col)
    fixed_headers(ws, ["#", "Key Activities & Scope", "Status", "SAP Start", "SAP End",
                       "APM Start", "APM End", "Dur (d)", "Mock Readiness", "Remark"], thin_gray)

    # ---- Gantt data rows: each activity = SAP band row + APM band row ----
    r = DATA_START_ROW
    activity_pairs = []           # (sap_row, apm_row) per activity
    for row in rows:
        if row["kind"] == "section":
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=last_col)
            c = ws.cell(row=r, column=1, value=row["label"])
            c.font = Font(bold=True, size=10, color="1F1F1F")
            c.alignment = Alignment(vertical="center")
            for col in range(1, last_col + 1):
                ws.cell(row=r, column=col).fill = solid_fill(SECTION_FILL)
            ws.row_dimensions[r].height = 18
            r += 1
            continue

        sap_row, apm_row = r, r + 1
        r += 2
        phase = row["phase"]
        cfg = PHASE[phase]

        # SAP band row (identity + SAP dates)
        status = row["status"] or ""
        a = ws.cell(row=sap_row, column=1, value=row["num"])
        a.alignment = Alignment(horizontal="center", vertical="center")
        nc = ws.cell(row=sap_row, column=2, value=row["name"])
        nc.alignment = Alignment(vertical="center", wrap_text=True)
        sc = ws.cell(row=sap_row, column=3, value=status)
        sc.alignment = Alignment(horizontal="center", vertical="center")
        if status.lower() == "completed":
            sc.font = Font(bold=True, color="375623", size=9)
        elif status.lower() == "in progress":
            sc.font = Font(bold=True, color="1F4E79", size=9)
        else:
            sc.font = Font(size=9, color="595959")

        if row["start"] and row["end"]:
            ws.cell(row=sap_row, column=4, value=row["start"]).number_format = "dd mmm yy"
            ws.cell(row=sap_row, column=5, value=row["end"]).number_format = "dd mmm yy"
        else:
            ws.cell(row=sap_row, column=4, value=row["timeline_raw"] or "TBC")
            ws.cell(row=sap_row, column=4).font = Font(italic=True, color="C00000", size=9)
            ws.cell(row=sap_row, column=5, value="")
        for col in (4, 5):
            ws.cell(row=sap_row, column=col).alignment = Alignment(horizontal="center")

        fc = ws.cell(row=sap_row, column=8,
                     value=f'=IF(AND(ISNUMBER($D{sap_row}),ISNUMBER($E{sap_row})),'
                           f'$E{sap_row}-$D{sap_row}+1,"")')
        fc.alignment = Alignment(horizontal="center")
        write_mock(ws, sap_row, row, col=9)
        remark = (row.get("remark") or "").strip()
        if remark:
            rk = ws.cell(row=sap_row, column=10, value=remark)
            rk.alignment = Alignment(wrap_text=True, vertical="top", horizontal="left")
            rk.font = Font(size=8, color="404040")
            lines = remark.count("\n") + 1
            if lines > 1:
                ws.row_dimensions[sap_row].height = max(20, 13 * lines + 6)

        bar_cf(ws, sap_row, mfc, last_col, f"D{sap_row}", f"E{sap_row}",
               cfg["solid"], cfg["border"])

        # APM band row
        if row["apm_start"] and row["apm_end"]:
            ws.cell(row=apm_row, column=6, value=row["apm_start"]).number_format = "dd mmm yy"
            ws.cell(row=apm_row, column=7, value=row["apm_end"]).number_format = "dd mmm yy"
        else:
            tbc = ws.cell(row=apm_row, column=6, value="TBC")
            tbc.font = Font(italic=True, bold=True, color="C00000", size=9)
            ws.cell(row=apm_row, column=7, value="")
        for col in (6, 7):
            ws.cell(row=apm_row, column=col).alignment = Alignment(horizontal="center")

        bar_cf(ws, apm_row, mfc, last_col, f"F{apm_row}", f"G{apm_row}",
               APM_FILL, APM_BORDER)

        tint_cf(ws, sap_row, apm_row, mfc, last_col)
        ws.row_dimensions[sap_row].height = 20
        ws.row_dimensions[apm_row].height = 13
        activity_pairs.append((sap_row, apm_row))
        row["sap_row"] = sap_row
        row["apm_row"] = apm_row

    mil_row = milestone_row(ws, r, mfc, last_col, GO_LIVE)

    # ---- FTE tracking section: position x month (no activity breakdown) ----
    groups = month_groups(weeks)     # same month boundaries as the Gantt header
    month_labels = [f"{SHORT_MONTH[gm]}'{str(gy)[2:]}" for _, _, gy, gm in groups]

    r = mil_row + 2
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=last_col)
    ft = ws.cell(row=r, column=1, value=("FTE TRACKING - monthly FTE by position "
                                         f"({month_labels[0]} to {month_labels[-1]})"))
    ft.fill = solid_fill("1F3864")
    ft.font = Font(bold=True, size=11, color="FFFFFF")
    ft.alignment = Alignment(vertical="center")
    ws.row_dimensions[r].height = 20
    r += 1
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=last_col)
    fs = ws.cell(row=r, column=1, value=(
        "1.0 = one person for the full month. One row per person (position repeated per person). "
        "Values are MOCK SAMPLES - type over them. The 'Total FTE' row auto-sums each month."))
    fs.fill = solid_fill(SUBTITLE_FILL)
    fs.font = Font(italic=True, size=9, color="404040")
    fs.alignment = Alignment(vertical="center", wrap_text=True)
    ws.row_dimensions[r].height = 18
    r += 1

    # header row: 'Position' + one merged month cell per month band
    hc = ws.cell(row=r, column=2, value="Position")
    hc.font = Font(bold=True, size=9)
    hc.fill = solid_fill(HEADER_FILL)
    hc.border = Border(left=thin_gray, right=thin_gray, top=thin_gray, bottom=thin_gray)
    for col in (1, 3, 4, 5, 6, 7, 8, 9, 10):
        ws.cell(row=r, column=col).fill = solid_fill(HEADER_FILL)
    for (gs, ge, gy, gm), label in zip(groups, month_labels):
        c0, c1 = mfc + gs, mfc + ge
        if c1 > c0:
            ws.merge_cells(start_row=r, start_column=c0, end_row=r, end_column=c1)
        mc = ws.cell(row=r, column=c0, value=label)
        mc.font = Font(bold=True, size=8)
        mc.alignment = Alignment(horizontal="center", vertical="center")
        mc.fill = solid_fill(HEADER_FILL)
        for col in range(c0, c1 + 1):
            ws.cell(row=r, column=col).border = Border(
                left=thin_gray, right=thin_gray, top=thin_gray, bottom=thin_gray)
    ws.row_dimensions[r].height = 14
    hdr_row = r
    r += 1

    # mock monthly values: Feb+Mar 26 = 0, Apr-Dec 26 + Jan-Sep 27 = sample, Oct-Dec 27 = 0
    def month_values(sample):
        return [0.0, 0.0] + sample + sample + [0.0, 0.0, 0.0]

    pos_first = r
    for name, sample in POSITIONS:
        vals = month_values(sample)
        nc = ws.cell(row=r, column=2, value=name)
        nc.font = Font(size=9, color="1F1F1F")
        nc.alignment = Alignment(vertical="center")
        for (gs, ge, _gy, _gm), v in zip(groups, vals):
            c0, c1 = mfc + gs, mfc + ge
            if c1 > c0:
                ws.merge_cells(start_row=r, start_column=c0, end_row=r, end_column=c1)
            cell = ws.cell(row=r, column=c0, value=v)
            cell.font = Font(size=9)
            cell.alignment = Alignment(horizontal="right", vertical="center")
            cell.number_format = "0.00"
            for col in range(c0, c1 + 1):
                ws.cell(row=r, column=col).border = Border(
                    left=thin_gray, right=thin_gray, top=thin_gray, bottom=thin_gray)
        ws.row_dimensions[r].height = 14
        r += 1
    pos_last = r - 1

    # total row: auto-sum of all positions for each month
    tot_row = r
    tc = ws.cell(row=tot_row, column=2, value="Total FTE (all positions)")
    tc.font = Font(bold=True, size=9, color="1F1F1F")
    tc.fill = solid_fill(FTE_SUBTOTAL_FILL)
    for (gs, ge, _gy, _gm) in groups:
        c0, c1 = mfc + gs, mfc + ge
        L = get_column_letter(c0)
        if c1 > c0:
            ws.merge_cells(start_row=tot_row, start_column=c0, end_row=tot_row, end_column=c1)
        tcell = ws.cell(row=tot_row, column=c0,
                        value=f"=SUM({L}{pos_first}:{L}{pos_last})")
        tcell.font = Font(bold=True, size=9)
        tcell.alignment = Alignment(horizontal="right", vertical="center")
        tcell.number_format = "0.00"
        tcell.fill = solid_fill(FTE_SUBTOTAL_FILL)
        for col in range(c0, c1 + 1):
            ws.cell(row=tot_row, column=col).fill = solid_fill(FTE_SUBTOTAL_FILL)
            ws.cell(row=tot_row, column=col).border = Border(
                left=thin_gray, right=thin_gray, top=thin_gray, bottom=thin_gray)
    ws.row_dimensions[tot_row].height = 14
    r += 1
    r += 1                                  # spacer before notes

    notes_r = r
    ws.cell(row=notes_r, column=1, value="Notes:").font = Font(bold=True)
    notes_r += 1
    for note in [
        "1. Each activity has TWO stacked bars: phase color = SAP window (SAP Start/End D:E); "
        "blue = APM's own window (APM Start/End F:G). Red 'TBC' = not confirmed yet - the blue "
        "bar appears once real dates are typed.",
        "2. FTE grid below: position x month, one row per person (mock samples - type over them). "
        "1.0 = one person for the full month. The 'Total FTE (all positions)' row auto-sums each "
        "month; monthly columns match the Gantt month bands (Feb'26 - Dec'27).",
        "3. C. Mock Validation uses the switch-source approach; APM needs 5-10 extra days after "
        "post-load (GCP team prepares legacy data).",
        "4. Many dates are TBC - waiting for confirmation from SHANA (SIT, UAT, Business Sim, "
        "Mock 3, Ramp Down, Blackout, Ramp Up, Hypercare).",
        "5. Bars are conditional formatting; Dur (d) and FTE Total are formulas - all update "
        "automatically. Go-Live date is editable in its milestone row.",
        "6. Status has a dropdown; SAP/APM date columns open a date picker.",
        "7. 'Mock Readiness' (col I): readiness dates for Mock 1-3 (from source col E); "
        "grey hatch = not required.",
        "8. APM dates seeded from source col F: #1 mirrors the SAP lifecycle; CIT and Mock 2 "
        "(9-Oct-26) parsed; every unconfirmed row is left as TBC - edit freely.",
        "9. 'SAP Timeline' tab is a read-only linked view for stakeholders who only need the "
        "SAP schedule - edit nothing there.",
        "10. Re-run build_gantt.py to regenerate from the source plan; previous versions are "
        "backed up automatically.",
    ]:
        c = ws.cell(row=notes_r, column=1, value=note)
        c.font = Font(size=9, color="404040")
        notes_r += 1

    # ---- data validation: status dropdown + date pickers ----
    status_dv = DataValidation(type="list", formula1='"Not started,In progress,Completed"',
                               allow_blank=True, showErrorMessage=True,
                               errorTitle="Invalid status",
                               error="Choose one of: Not started, In progress, Completed",
                               showInputMessage=True, promptTitle="Status",
                               prompt="Select from the dropdown")
    ws.add_data_validation(status_dv)
    date_dv = DataValidation(type="date", operator="between",
                             formula1="DATE(2026,1,1)", formula2="DATE(2030,12,31)",
                             allow_blank=True, showErrorMessage=True,
                             errorTitle="Invalid date",
                             error="Pick a date between 2026 and 2030 (use the calendar picker)",
                             showInputMessage=True, promptTitle="Start / End date",
                             prompt="Click to open the date picker")
    ws.add_data_validation(date_dv)
    for sap_row, apm_row in activity_pairs:
        status_dv.add(f"C{sap_row}")
        date_dv.add(f"D{sap_row}:E{sap_row}")
        date_dv.add(f"F{apm_row}:G{apm_row}")
    date_dv.add(f"D{mil_row}")

    today_header_cf(ws, mfc, last_col)
    ws.freeze_panes = f"{first_letter}6"
    return {"mil_row": mil_row, "pairs": activity_pairs}


# =====================================================================
#  TAB 1 (view): 'SAP Timeline'
# =====================================================================
def build_sap_view(wb, rows, weeks, master):
    ws = wb.create_sheet("SAP Timeline")
    ws.sheet_view.showGridLines = False

    n = len(weeks)
    mfc = 9                        # fixed block A..H, weeks start col I
    first_letter = get_column_letter(mfc)
    last_col = (mfc - 1) + n
    last_letter = get_column_letter(last_col)
    mname = "'APM Plan & FTE'"

    for c, wd in {1: 4.5, 2: 62, 3: 12, 4: 13, 5: 13, 6: 10, 7: 13, 8: 45}.items():
        ws.column_dimensions[get_column_letter(c)].width = wd
    for i in range(n):
        ws.column_dimensions[get_column_letter(mfc + i)].width = VIEW_WEEK_WIDTH

    titles(ws, last_col, "APM - SAP Activity & Support Plan  |  SAP Timeline (Weekly)",
           "SAP schedule view for SAP-only stakeholders - read-only: dates/status LINK from the "
           "'APM Plan & FTE' tab (edit them there, this view updates automatically) | "
           "orange header + gray band = current week; red block = Go-Live week | "
           "AS OF 27 AUG 2026 - SAP Go-Live 1 Jul 2027")

    # legend (row 3)
    leg = ws.cell(row=3, column=1, value="Phase:")
    leg.font = Font(bold=True, size=9)
    leg.alignment = Alignment(horizontal="right", vertical="center")
    thin_dark = Side(style="thin", color="404040")
    for i, letter in enumerate(PHASE):
        chip = ws.cell(row=3, column=2 + i, value=letter)
        chip.fill = solid_fill(PHASE[letter]["solid"])
        chip.font = Font(bold=True, size=9, color="1F1F1F")
        chip.alignment = Alignment(horizontal="center", vertical="center")
        chip.border = Border(left=thin_dark, right=thin_dark, top=thin_dark, bottom=thin_dark)
    hatch = ws.cell(row=3, column=7, value="")
    hatch.fill = PatternFill(patternType="lightUp", fgColor="BFBFBF", bgColor="FFFFFF")
    hatch.border = Border(left=thin_dark, right=thin_dark, top=thin_dark, bottom=thin_dark)
    hl = ws.cell(row=3, column=8, value="No required info")
    hl.font = Font(italic=True, size=8, color="595959")
    hl.alignment = Alignment(vertical="center")
    ws.row_dimensions[3].height = 14

    thin_gray = week_headers(ws, weeks, mfc, last_col)
    fixed_headers(ws, ["#", "Key Activities & Scope", "Status", "SAP Start", "SAP End",
                       "Dur (d)", "Mock Readiness", "Remark"], thin_gray)

    # map activity number -> master (sap row)
    num_to_sap = {i + 1: pair[0] for i, pair in enumerate(master["pairs"])}

    # ---- data rows (one row per activity; dates/status linked to master) ----
    r = DATA_START_ROW
    for row in rows:
        if row["kind"] == "section":
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=last_col)
            c = ws.cell(row=r, column=1, value=row["label"])
            c.font = Font(bold=True, size=10, color="1F1F1F")
            c.alignment = Alignment(vertical="center")
            for col in range(1, last_col + 1):
                ws.cell(row=r, column=col).fill = solid_fill(SECTION_FILL)
            ws.row_dimensions[r].height = 18
            r += 1
            continue

        msap = num_to_sap[row["num"]]
        cfg = PHASE[row["phase"]]
        a = ws.cell(row=r, column=1, value=row["num"])
        a.alignment = Alignment(horizontal="center", vertical="center")
        nc = ws.cell(row=r, column=2, value=row["name"])
        nc.alignment = Alignment(vertical="center", wrap_text=True)
        sc = ws.cell(row=r, column=3, value=f"={mname}!C{msap}")
        sc.alignment = Alignment(horizontal="center", vertical="center")
        sc.font = Font(bold=True, size=9, color="1F4E79")

        d1 = ws.cell(row=r, column=4, value=f"={mname}!D{msap}")
        d1.number_format = "dd mmm yy"
        d2 = ws.cell(row=r, column=5, value=f"={mname}!E{msap}")
        d2.number_format = "dd mmm yy"
        for col in (4, 5):
            ws.cell(row=r, column=col).alignment = Alignment(horizontal="center")

        fc = ws.cell(row=r, column=6,
                     value=f'=IF(AND(ISNUMBER($D{r}),ISNUMBER($E{r})),$E{r}-$D{r}+1,"")')
        fc.alignment = Alignment(horizontal="center")
        write_mock(ws, r, row, col=7)
        remark = (row.get("remark") or "").strip()
        if remark:
            rk = ws.cell(row=r, column=8, value=remark)
            rk.alignment = Alignment(wrap_text=True, vertical="top", horizontal="left")
            rk.font = Font(size=8, color="404040")
            lines = remark.count("\n") + 1
            if lines > 1:
                ws.row_dimensions[r].height = max(20, 13 * lines + 6)

        bar_cf(ws, r, mfc, last_col, f"D{r}", f"E{r}", cfg["solid"], cfg["border"])
        tint_cf(ws, r, r, mfc, last_col)
        ws.row_dimensions[r].height = 20
        r += 1

    # milestone row: date linked to master milestone row
    mrow = r
    ws.merge_cells(start_row=mrow, start_column=1, end_row=mrow, end_column=2)
    mc = ws.cell(row=mrow, column=1, value="SAP Go-Live (Milestone)")
    mc.font = Font(bold=True, size=10, color="FFFFFF")
    mc.alignment = Alignment(vertical="center")
    for col in range(1, last_col + 1):
        ws.cell(row=mrow, column=col).fill = solid_fill("404040")
    gc = ws.cell(row=mrow, column=4, value=f"={mname}!D{master['mil_row']}")
    gc.number_format = "dd mmm yy"
    gc.alignment = Alignment(horizontal="center")
    gc.font = Font(color="FFFFFF")
    c = first_letter
    red_side = Side(style="thin", color="7F0000")
    dxf = DifferentialStyle(
        fill=PatternFill(start_color=MILESTONE_FILL, end_color=MILESTONE_FILL,
                         fill_type="solid"),
        border=Border(left=red_side, right=red_side, top=red_side, bottom=red_side))
    rule = Rule(type="expression", dxf=dxf)
    rule.formula = [f"AND({c}$5<=$D{mrow},{c}$5+6>=$D{mrow})"]
    ws.conditional_formatting.add(f"{c}{mrow}:{last_letter}{mrow}", rule)
    ws.row_dimensions[mrow].height = 18

    today_header_cf(ws, mfc, last_col)

    # footer note
    r = mrow + 2
    ws.cell(row=r, column=1, value="Notes:").font = Font(bold=True)
    r += 1
    for note in [
        "1. READ-ONLY SAP schedule view. Dates, status and the Go-Live date are linked formulas "
        "from the 'APM Plan & FTE' tab - edit them there, this view updates automatically.",
        "2. One row per activity = the SAP window (phase color). APM's own dates and FTE "
        "planning live on the 'APM Plan & FTE' tab.",
        "3. Many dates are TBC - waiting for confirmation from SHANA (SIT, UAT, Business Sim, "
        "Mock 3, Ramp Down, Blackout, Ramp Up, Hypercare).",
        "4. Orange header + gray band = current week; red block = SAP Go-Live week (1 Jul 2027).",
    ]:
        c = ws.cell(row=r, column=1, value=note)
        c.font = Font(size=9, color="404040")
        r += 1

    ws.freeze_panes = f"{first_letter}6"


# =====================================================================
def build():
    backup_previous()
    rows = load_rows()
    weeks = calc_weeks()

    wb = openpyxl.Workbook()
    master = build_master(wb, rows, weeks)
    build_sap_view(wb, rows, weeks, master)
    # put the SAP view first
    wb.move_sheet("SAP Timeline", offset=-1)
    wb.save(OUT)

    print(f"Saved: {OUT}")
    print(f"Sheets: {wb.sheetnames}")
    print(f"Week grid: {weeks[0]} .. {weeks[-1]} ({len(weeks)} week columns)")
    for row in rows:
        if row["kind"] == "section":
            continue
        sd = row["start"].isoformat() if row["start"] else "TBC"
        ad = row["apm_start"].isoformat() if row["apm_start"] else "TBC"
        print(f"  #{row['num']:>2} {row['name'][:46]:<46} SAP {sd:<12} "
              f"APM {ad:<12} window: {window_text(row)}")


if __name__ == "__main__":
    build()
