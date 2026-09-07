"""
Add Gantt Chart / Sprint Timeline Visualization to the Summary sheet.
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from collections import defaultdict, Counter

FILE = r"exel file/Project Information 2.xlsx"

# ── Colours ─────────────────────────────────────────────────────────────────
DARK_BLUE  = "1F3864"
MED_BLUE   = "2E75B6"
LIGHT_BLUE = "D6E4F0"
WHITE      = "FFFFFF"
GRAY_BG    = "F2F2F2"
GRAY2      = "D9D9D9"

# Module bar colours (pastel / readable on screen)
MOD_COLORS = {
    "asm":           "4472C4",   # blue
    "rca":           "70AD47",   # green
    "ui improvement":"ED7D31",   # orange
    "handover":      "A5A5A5",   # gray
    "sap":           "FFC000",   # gold / yellow
}
MOD_FONT_COLORS = {
    "asm":           "FFFFFF",
    "rca":           "FFFFFF",
    "ui improvement":"FFFFFF",
    "handover":      "FFFFFF",
    "sap":           "000000",
}

thin_border = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"), bottom=Side(style="thin"),
)
center = Alignment(horizontal="center", vertical="center", wrap_text=True)
left_wrap = Alignment(horizontal="left", vertical="center", wrap_text=True)

hdr_font   = Font(name="Calibri", bold=True, color=WHITE, size=9)
hdr_fill   = PatternFill("solid", fgColor=DARK_BLUE)
sub_font   = Font(name="Calibri", bold=True, color=DARK_BLUE, size=10)
sub_fill   = PatternFill("solid", fgColor=LIGHT_BLUE)
section_font = Font(name="Calibri", bold=True, color=WHITE, size=11)
section_fill = PatternFill("solid", fgColor=MED_BLUE)
bold9      = Font(name="Calibri", bold=True, size=9)
normal9    = Font(name="Calibri", size=9)


def section_header(ws, row, text, cols):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=cols)
    cell = ws.cell(row=row, column=1, value=text)
    cell.font = section_font
    cell.fill = section_fill
    cell.alignment = Alignment(horizontal="left", vertical="center")
    for c in range(1, cols + 1):
        ws.cell(row=row, column=c).fill = section_fill
        ws.cell(row=row, column=c).border = thin_border


def main():
    wb = openpyxl.load_workbook(FILE)
    ws_bl = wb["baseline"]
    ws = wb["Summary"]

    # ── Read baseline data ────────────────────────────────────────────────────
    rows_data = []
    for row in ws_bl.iter_rows(min_row=2, max_row=ws_bl.max_row, min_col=1, max_col=11, values_only=True):
        if row[0] is not None or row[2] is not None:
            rows_data.append({
                "module": row[0],
                "sprint": row[4],
                "status": row[7],
                "is_baseline": row[8],
            })

    # Sprint dates from sprintday sheet
    ws_sd = wb["sprintday"]
    sprint_dates = {}
    for row in ws_sd.iter_rows(min_row=6, max_row=ws_sd.max_row, min_col=1, max_col=2, values_only=True):
        if row[0] and row[1]:
            sprint_dates[row[0]] = row[1]

    # Module -> set of sprints with activity
    mod_sprints = defaultdict(set)
    for r in rows_data:
        if r["sprint"]:
            mod_sprints[r["module"]].add(r["sprint"])

    # Sprint -> module item count
    sprint_mod_counts = defaultdict(lambda: defaultdict(int))
    for r in rows_data:
        if r["sprint"]:
            sprint_mod_counts[r["sprint"]][r["module"]] += 1

    sprint_counts = Counter(r["sprint"] for r in rows_data if r["sprint"])

    # Module display order
    module_order = ["asm", "sap", "rca", "ui improvement", "handover"]
    sprint_range = list(range(18, 44))  # Sprint 18-43
    num_sprints = len(sprint_range)

    # ── Find where to append (after existing content) ─────────────────────────
    start_row = ws.max_row + 2

    # ── GANTT SECTION 6 ──────────────────────────────────────────────────────
    # Total columns needed: label (col 1) + 26 sprints = 27 cols + 1 for total
    # We'll use cols A (label) through AB (total) = 28 cols
    total_cols = 1 + num_sprints + 1  # label + sprints + total

    section_header(ws, start_row,
                   "6  |  GANTT CHART – Sprint Timeline by Module", total_cols)
    start_row += 1

    # Sub-header
    ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=total_cols)
    ws.cell(row=start_row, column=1,
            value="Color bars indicate which modules are active in each sprint. "
                  "Yellow = SAP activity (ECC to S/4HANA). Current date: 24 Aug 2026 (approx Sprint 34)."
           ).font = Font(name="Calibri", size=9, italic=True, color="555555")
    start_row += 1

    # ── Row: Month labels ─────────────────────────────────────────────────────
    r = start_row
    ws.cell(row=r, column=1, value="Module").font = hdr_font
    ws.cell(row=r, column=1).fill = hdr_fill
    ws.cell(row=r, column=1).alignment = center
    ws.cell(row=r, column=1).border = thin_border

    # Sprint header row
    r_sprint = r
    for i, sp in enumerate(sprint_range):
        col = 2 + i
        ws.cell(row=r, column=col, value=f"S{sp}").font = hdr_font
        ws.cell(row=r, column=col).fill = hdr_fill
        ws.cell(row=r, column=col).alignment = center
        ws.cell(row=r, column=col).border = thin_border
        ws.column_dimensions[get_column_letter(col)].width = 4.5

    # Total column
    ws.cell(row=r, column=total_cols, value="TOTAL").font = hdr_font
    ws.cell(row=r, column=total_cols).fill = hdr_fill
    ws.cell(row=r, column=total_cols).alignment = center
    ws.cell(row=r, column=total_cols).border = thin_border
    ws.column_dimensions[get_column_letter(1)].width = 18
    ws.column_dimensions[get_column_letter(total_cols)].width = 7
    start_row += 1

    # ── Row: Date labels ──────────────────────────────────────────────────────
    r_date = start_row
    ws.cell(row=r_date, column=1, value="Start Date").font = bold9
    ws.cell(row=r_date, column=1).fill = sub_fill
    ws.cell(row=r_date, column=1).alignment = center
    ws.cell(row=r_date, column=1).border = thin_border

    for i, sp in enumerate(sprint_range):
        col = 2 + i
        dt = sprint_dates.get(sp, "")
        if dt:
            val = dt.strftime("%d-%b") if hasattr(dt, "strftime") else str(dt)
        else:
            val = ""
        ws.cell(row=r_date, column=col, value=val).font = Font(name="Calibri", size=7, color="333333")
        ws.cell(row=r_date, column=col).fill = sub_fill
        ws.cell(row=r_date, column=col).alignment = center
        ws.cell(row=r_date, column=col).border = thin_border

    ws.cell(row=r_date, column=total_cols).fill = sub_fill
    ws.cell(row=r_date, column=total_cols).border = thin_border
    start_row += 1

    # ── Module Gantt rows ─────────────────────────────────────────────────────
    for mod in module_order:
        r = start_row
        ws.cell(row=r, column=1, value=mod.upper()).font = bold9
        ws.cell(row=r, column=1).alignment = Alignment(horizontal="left", vertical="center")
        ws.cell(row=r, column=1).border = thin_border

        mod_total = 0
        for i, sp in enumerate(sprint_range):
            col = 2 + i
            count = sprint_mod_counts[sp].get(mod, 0)
            mod_total += count

            if count > 0:
                fill_color = MOD_COLORS.get(mod, "999999")
                font_color = MOD_FONT_COLORS.get(mod, "FFFFFF")
                ws.cell(row=r, column=col, value=str(count)).font = Font(
                    name="Calibri", bold=True, size=9, color=font_color
                )
                ws.cell(row=r, column=col).fill = PatternFill("solid", fgColor=fill_color)
                # SAP gets an extra border highlight
                if mod == "sap":
                    ws.cell(row=r, column=col).border = Border(
                        left=Side(style="thin"), right=Side(style="thin"),
                        top=Side(style="medium", color="BF8F00"),
                        bottom=Side(style="medium", color="BF8F00"),
                    )
            else:
                ws.cell(row=r, column=col, value="").font = normal9
                ws.cell(row=r, column=col).fill = PatternFill("solid", fgColor="F5F5F5")
                ws.cell(row=r, column=col).border = thin_border

            ws.cell(row=r, column=col).alignment = center

        # Total column
        ws.cell(row=r, column=total_cols, value=mod_total).font = bold9
        ws.cell(row=r, column=total_cols).alignment = center
        ws.cell(row=r, column=total_cols).border = thin_border
        ws.cell(row=r, column=total_cols).fill = PatternFill("solid", fgColor=GRAY_BG)

        start_row += 1

    # ── Sprint Total row ──────────────────────────────────────────────────────
    r = start_row
    ws.cell(row=r, column=1, value="SPRINT TOTAL").font = Font(name="Calibri", bold=True, size=9, color=DARK_BLUE)
    ws.cell(row=r, column=1).alignment = Alignment(horizontal="left", vertical="center")
    ws.cell(row=r, column=1).border = thin_border
    ws.cell(row=r, column=1).fill = PatternFill("solid", fgColor=GRAY_BG)

    grand_total = 0
    for i, sp in enumerate(sprint_range):
        col = 2 + i
        count = sprint_counts.get(sp, 0)
        grand_total += count
        ws.cell(row=r, column=col, value=count).font = Font(name="Calibri", bold=True, size=9)
        ws.cell(row=r, column=col).alignment = center
        ws.cell(row=r, column=col).border = thin_border
        # Color intensity based on count
        if count >= 10:
            ws.cell(row=r, column=col).fill = PatternFill("solid", fgColor="B4C6E7")
        elif count >= 5:
            ws.cell(row=r, column=col).fill = PatternFill("solid", fgColor="D6E4F0")
        else:
            ws.cell(row=r, column=col).fill = PatternFill("solid", fgColor="E9EFF7")

    ws.cell(row=r, column=total_cols, value=grand_total).font = Font(name="Calibri", bold=True, size=9)
    ws.cell(row=r, column=total_cols).alignment = center
    ws.cell(row=r, column=total_cols).border = thin_border
    ws.cell(row=r, column=total_cols).fill = PatternFill("solid", fgColor=GRAY2)
    start_row += 1

    # ── Legend row ─────────────────────────────────────────────────────────────
    r = start_row
    ws.cell(row=r, column=1, value="LEGEND").font = Font(name="Calibri", bold=True, size=8, color=DARK_BLUE)
    ws.cell(row=r, column=1).border = thin_border
    legend_items = [
        ("ASM", MOD_COLORS["asm"], MOD_FONT_COLORS["asm"]),
        ("RCA", MOD_COLORS["rca"], MOD_FONT_COLORS["rca"]),
        ("UI IMP", MOD_COLORS["ui improvement"], MOD_FONT_COLORS["ui improvement"]),
        ("HANDOVER", MOD_COLORS["handover"], MOD_FONT_COLORS["handover"]),
        ("SAP", MOD_COLORS["sap"], MOD_FONT_COLORS["sap"]),
    ]
    col = 2
    for label, bg, fg in legend_items:
        ws.cell(row=r, column=col, value=label).font = Font(name="Calibri", bold=True, size=8, color=fg)
        ws.cell(row=r, column=col).fill = PatternFill("solid", fgColor=bg)
        ws.cell(row=r, column=col).alignment = center
        ws.cell(row=r, column=col).border = thin_border
        col += 1
        ws.cell(row=r, column=col, value="").border = thin_border
        col += 1

    start_row += 2

    # ══════════════════════════════════════════════════════════════════════════
    # SECTION 7  –  Sprint-by-Sprint Detail Table
    # ══════════════════════════════════════════════════════════════════════════
    section_header(ws, start_row,
                   "7  |  SPRINT-BY-SPRINT BREAKDOWN", 8)
    start_row += 1

    headers = ["Sprint", "Date Range", "ASM", "SAP", "RCA", "UI IMP", "HANDOVER", "Total"]
    for i, h in enumerate(headers):
        c = ws.cell(row=start_row, column=1 + i, value=h)
        c.font = hdr_font
        c.fill = hdr_fill
        c.alignment = center
        c.border = thin_border
    start_row += 1

    for sp in sprint_range:
        r = start_row
        ws.cell(row=r, column=1, value=f"Sprint {sp}").font = bold9
        ws.cell(row=r, column=1).border = thin_border
        ws.cell(row=r, column=1).alignment = center

        # Date range
        dt = sprint_dates.get(sp, "")
        if dt:
            val = dt.strftime("%d %b %Y") if hasattr(dt, "strftime") else str(dt)
        else:
            val = ""
        ws.cell(row=r, column=2, value=val).font = normal9
        ws.cell(row=r, column=2).border = thin_border
        ws.cell(row=r, column=2).alignment = center

        row_total = 0
        for j, mod in enumerate(["asm", "sap", "rca", "ui improvement", "handover"]):
            col = 3 + j
            count = sprint_mod_counts[sp].get(mod, 0)
            row_total += count
            ws.cell(row=r, column=col, value=count if count > 0 else "").font = normal9
            ws.cell(row=r, column=col).alignment = center
            ws.cell(row=r, column=col).border = thin_border
            if count > 0:
                ws.cell(row=r, column=col).fill = PatternFill("solid", fgColor=MOD_COLORS[mod])
                ws.cell(row=r, column=col).font = Font(name="Calibri", bold=True, size=9,
                                                        color=MOD_FONT_COLORS[mod])

        ws.cell(row=r, column=8, value=row_total).font = bold9
        ws.cell(row=r, column=8).alignment = center
        ws.cell(row=r, column=8).border = thin_border
        ws.cell(row=r, column=8).fill = PatternFill("solid", fgColor=GRAY_BG)

        # Alternate row shading
        if (sp - 18) % 2 == 0:
            for col_idx in [1, 2]:
                ws.cell(row=r, column=col_idx).fill = PatternFill("solid", fgColor="E9EFF7")

        start_row += 1

    # Grand total
    r = start_row
    ws.cell(row=r, column=1, value="TOTAL").font = Font(name="Calibri", bold=True, size=9, color=DARK_BLUE)
    ws.cell(row=r, column=1).fill = PatternFill("solid", fgColor=GRAY_BG)
    ws.cell(row=r, column=1).border = thin_border
    ws.cell(row=r, column=2).fill = PatternFill("solid", fgColor=GRAY_BG)
    ws.cell(row=r, column=2).border = thin_border

    grand = 0
    for j, mod in enumerate(module_order):
        col = 3 + j
        total = sum(sprint_mod_counts[sp].get(mod, 0) for sp in sprint_range)
        grand += total
        ws.cell(row=r, column=col, value=total).font = bold9
        ws.cell(row=r, column=col).alignment = center
        ws.cell(row=r, column=col).border = thin_border
        ws.cell(row=r, column=col).fill = PatternFill("solid", fgColor=GRAY_BG)

    ws.cell(row=r, column=8, value=grand).font = Font(name="Calibri", bold=True, size=10, color=DARK_BLUE)
    ws.cell(row=r, column=8).alignment = center
    ws.cell(row=r, column=8).border = thin_border
    ws.cell(row=r, column=8).fill = PatternFill("solid", fgColor=GRAY2)

    # ── Set column widths for detail table ─────────────────────────────────────
    ws.column_dimensions[get_column_letter(2)].width = 14
    for j in range(3, 8):
        ws.column_dimensions[get_column_letter(j)].width = 9

    # ── Freeze panes & print setup ────────────────────────────────────────────
    ws.sheet_view.showGridLines = False

    wb.save(FILE)
    print("OK  Gantt chart and Sprint breakdown added to Summary sheet")


if __name__ == "__main__":
    main()
