"""
Build a proper Gantt chart in the APM WBS Schedule sheet.

Replaces the old monthly ■-character Gantt with a professional weekly-column
Gantt that uses blue cell fills, monthly group headers, and weekly sub-headers.

Visual style matches the reference picture:
- Row 3: Month headers (merged across their week columns)
- Row 4: Week numbers (W1, W2, ...) + day-of-month markers
- Rows 5-36: Task data with blue filled bars in the Gantt columns
"""

import sys
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

import calendar
from datetime import date, timedelta

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

FILE_IN  = r"../APM SAP Plan (excel)/apm-wbs-template-v4-fixed.xlsx"
FILE_OUT = r"../APM SAP Plan (excel)/apm-wbs-template-v4.xlsx"

# ── Colours ─────────────────────────────────────────────────────────────────
BAR_BLUE   = "4472C4"   # Gantt bar fill
DARK_BLUE  = "1F3864"
MED_BLUE   = "2E75B6"
LIGHT_BLUE = "D6E4F0"
WHITE      = "FFFFFF"
GRAY_BG    = "F2F2F2"
TODAY_CLR  = "FFC000"   # Yellow highlight for today column

# ── Styles ──────────────────────────────────────────────────────────────────
bar_fill = PatternFill("solid", fgColor=BAR_BLUE)
empty_fill = PatternFill("solid", fgColor="F8F8F8")
hdr_fill = PatternFill("solid", fgColor=DARK_BLUE)
month_fill = PatternFill("solid", fgColor=MED_BLUE)
sub_fill = PatternFill("solid", fgColor=LIGHT_BLUE)
today_fill = PatternFill("solid", fgColor="FFFFCC")

hdr_font = Font(name="Calibri", bold=True, color=WHITE, size=9)
month_font = Font(name="Calibri", bold=True, color=WHITE, size=9)
week_font = Font(name="Calibri", size=7, color="333333")
day_font = Font(name="Calibri", size=6, color="666666")
normal_font = Font(name="Calibri", size=9)
bold_font = Font(name="Calibri", bold=True, size=9)

thin_border = Border(
    left=Side(style="thin", color="D0D0D0"),
    right=Side(style="thin", color="D0D0D0"),
    top=Side(style="thin", color="D0D0D0"),
    bottom=Side(style="thin", color="D0D0D0"),
)
center = Alignment(horizontal="center", vertical="center", wrap_text=False)
left_center = Alignment(horizontal="left", vertical="center")

# ── Timeline definition ─────────────────────────────────────────────────────
MONTHS = [
    (2026, 8), (2026, 9), (2026, 10), (2026, 11), (2026, 12),
    (2027, 1), (2027, 2), (2027, 3), (2027, 4), (2027, 5), (2027, 6), (2027, 7),
]

TODAY = date(2026, 8, 26)


def get_month_weeks(year, month):
    """Return list of (week_start, week_end) for each week in a month.
    Weeks are month-contained: W1=1-7, W2=8-14, W3=15-21, W4=22-28, W5=29-end."""
    _, last_day_num = calendar.monthrange(year, month)

    weeks = []
    day = 1
    while day <= last_day_num:
        week_start = date(year, month, day)
        week_end = date(year, month, min(day + 6, last_day_num))
        weeks.append((week_start, week_end))
        day += 7
    return weeks


def build_timeline():
    """Build the full timeline: list of (month_label, week_num_in_month, week_start, week_end)."""
    timeline = []
    for year, month in MONTHS:
        month_label = date(year, month, 1).strftime("%b '%y")
        weeks = get_month_weeks(year, month)
        for i, (ws, we) in enumerate(weeks):
            timeline.append((month_label, i + 1, ws, we))
    return timeline


def main():
    wb = openpyxl.load_workbook(FILE_IN)
    ws = wb["APM WBS Schedule"]

    timeline = build_timeline()
    num_weeks = len(timeline)
    print(f"Timeline: {num_weeks} weeks across {len(MONTHS)} months")
    for ml, wi, w_start, w_end in timeline:
        print(f"  {ml} W{wi}: {w_start} - {w_end}")

    # The Gantt columns start at column M (col 13)
    GANTT_START_COL = 13  # Column M
    GANTT_END_COL = GANTT_START_COL + num_weeks - 1

    # ── Step 1: Unmerge ALL cells first ──────────────────────────────────────
    merged = list(ws.merged_cells.ranges)
    for mc in merged:
        ws.unmerge_cells(str(mc))

    # ── Step 2: Clear old Gantt columns (M through X = cols 13-24) ──────────
    for r in range(1, ws.max_row + 1):
        for col in range(GANTT_START_COL, 25):  # Clear M-X
            cell = ws.cell(row=r, column=col)
            cell.value = None
            cell.fill = PatternFill()
            cell.font = Font()
            cell.border = Border()
            cell.alignment = Alignment()

    # ── Step 3: Fix row 1 title to span full width ───────────────────────────
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=GANTT_START_COL + num_weeks)

    # ── Step 3: Row 3 = Month headers (merged across week columns) ───────────
    MONTH_ROW = 3
    WEEK_ROW = 4
    DATA_START_ROW = 5

    # Group timeline entries by month label
    month_groups = []
    current_label = None
    current_start = None
    for i, (ml, wi, ws_date, we_date) in enumerate(timeline):
        if ml != current_label:
            if current_label is not None:
                month_groups.append((current_label, current_start, i - 1))
            current_label = ml
            current_start = i
    if current_label is not None:
        month_groups.append((current_label, current_start, len(timeline) - 1))

    print(f"\nMonth groups: {len(month_groups)}")
    for label, start_idx, end_idx in month_groups:
        start_col = GANTT_START_COL + start_idx
        end_col = GANTT_START_COL + end_idx
        print(f"  {label}: cols {get_column_letter(start_col)}-{get_column_letter(end_col)} "
              f"({end_idx - start_idx + 1} weeks)")

    # Write month headers in row 3
    for label, start_idx, end_idx in month_groups:
        start_col = GANTT_START_COL + start_idx
        end_col = GANTT_START_COL + end_idx
        if start_col != end_col:
            ws.merge_cells(start_row=MONTH_ROW, start_column=start_col,
                           end_row=MONTH_ROW, end_column=end_col)
        cell = ws.cell(row=MONTH_ROW, column=start_col, value=label)
        cell.font = month_font
        cell.fill = month_fill
        cell.alignment = center
        cell.border = thin_border
        # Fill all merged cells with same styling
        for c in range(start_col, end_col + 1):
            ws.cell(row=MONTH_ROW, column=c).fill = month_fill
            ws.cell(row=MONTH_ROW, column=c).border = thin_border

    # ── Step 4: Row 4 = Week number sub-headers ──────────────────────────────
    for i, (ml, wi, ws_date, we_date) in enumerate(timeline):
        col = GANTT_START_COL + i
        # Show week number
        cell = ws.cell(row=WEEK_ROW, column=col, value=f"W{wi}")
        cell.font = week_font
        cell.fill = sub_fill
        cell.alignment = center
        cell.border = thin_border
        # Show first day of month as day number on the first week
        if wi == 1:
            day_str = str(ws_date.day)
            # We'll add day markers in a helper approach: just show week for now

    # ── Step 5: Add day-of-month markers in a thin row between week header and data ──
    # Actually, let's add day numbers within each week column
    # For each week, show the starting day number
    DAY_ROW = WEEK_ROW  # We'll embed day info in the week row itself

    # ── Step 6: Fill Gantt bars for each task ────────────────────────────────
    for row_idx in range(DATA_START_ROW, ws.max_row + 1):
        # Read task start/end dates from columns E and F
        start_cell = ws.cell(row=row_idx, column=5)  # E
        end_cell = ws.cell(row=row_idx, column=6)     # F

        task_start = start_cell.value
        task_end = end_cell.value

        if task_start is None or task_end is None:
            continue

        # Convert to date if datetime
        if hasattr(task_start, "date"):
            task_start = task_start.date()
        if hasattr(task_end, "date"):
            task_end = task_end.date()

        if not isinstance(task_start, date) or not isinstance(task_end, date):
            continue

        # Fill cells for each week that overlaps with the task
        for i, (ml, wi, wk_start_date, wk_end_date) in enumerate(timeline):
            col = GANTT_START_COL + i
            cell = ws.cell(row=row_idx, column=col)

            # Check if task overlaps this week
            # Task spans [task_start, task_end], week spans [wk_start, wk_end]
            if task_start <= wk_end_date and task_end >= wk_start_date:
                cell.value = ""
                cell.fill = bar_fill
                # Make bar columns narrow
            else:
                cell.fill = empty_fill

            cell.border = thin_border
            cell.alignment = center

    # ── Step 7: Set column widths ────────────────────────────────────────────
    # Gantt columns: narrow
    for i in range(num_weeks):
        col = GANTT_START_COL + i
        ws.column_dimensions[get_column_letter(col)].width = 3.0

    # ── Step 8: Highlight today column ───────────────────────────────────────
    for i, (ml, wi, wk_start, wk_end) in enumerate(timeline):
        if wk_start <= TODAY <= wk_end:
            col = GANTT_START_COL + i
            # Highlight the header
            ws.cell(row=MONTH_ROW, column=col).fill = PatternFill("solid", fgColor=TODAY_CLR)
            ws.cell(row=WEEK_ROW, column=col).fill = PatternFill("solid", fgColor=TODAY_CLR)
            # Add a thin vertical marker
            for r in range(DATA_START_ROW, ws.max_row + 1):
                cell = ws.cell(row=r, column=col)
                if cell.fill == bar_fill:
                    # Task is active this week AND it's the current week
                    cell.border = Border(
                        left=Side(style="medium", color=TODAY_CLR),
                        right=Side(style="medium", color=TODAY_CLR),
                        top=Side(style="thin", color="D0D0D0"),
                        bottom=Side(style="thin", color="D0D0D0"),
                    )
            print(f"\nToday marker: {ml} W{wi} (col {get_column_letter(col)})")
            break

    # ── Step 9: Add a "Year 2026 / Year 2027" label above months ────────────
    # Check if we need a year header row above the month row
    # We'll add it as row 2.5 by inserting a row... actually let's skip for now
    # and just put year info in the month labels

    # ── Step 10: Freeze panes ────────────────────────────────────────────────
    ws.freeze_panes = f"{get_column_letter(GANTT_START_COL)}{DATA_START_ROW}"

    # ── Step 11: Reformat header row (row 4 original) ────────────────────────
    # The original row 4 had "Aug '26", "Sep '26" etc. Now it's the week row.
    # Make sure the data headers in A4:L4 still look good
    for col in range(1, GANTT_START_COL):
        cell = ws.cell(row=MONTH_ROW, column=col)
        if cell.value is None or cell.value == "":
            # Fill empty header cells in month row with dark blue
            cell.fill = hdr_fill
            cell.border = thin_border
        cell = ws.cell(row=WEEK_ROW, column=col)
        if cell.value is None or cell.value == "":
            cell.fill = sub_fill
            cell.border = thin_border

    # ── Save ─────────────────────────────────────────────────────────────────
    wb.save(FILE_OUT)
    print(f"\n✅ Saved to {FILE_OUT}")
    print(f"   Gantt: {num_weeks} weekly columns ({get_column_letter(GANTT_START_COL)}-{get_column_letter(GANTT_END_COL)})")
    print(f"   Tasks: rows {DATA_START_ROW}-{ws.max_row}")


if __name__ == "__main__":
    main()
