"""
Add Excel charts to the Summary sheet:
1. Clustered bar chart: Baseline vs Current per module
2. Stacked horizontal bar chart: Change type by module + data table
"""

import openpyxl
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.series import DataPoint
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

FILE = r"exel file/Project Information 2.xlsx"

# ── Styles ──────────────────────────────────────────────────────────────────
DARK_BLUE = "1F3864"
MED_BLUE  = "2E75B6"
WHITE     = "FFFFFF"
GRAY_BG   = "F2F2F2"

thin_border = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"), bottom=Side(style="thin"),
)
center = Alignment(horizontal="center", vertical="center", wrap_text=True)

hdr_font = Font(name="Calibri", bold=True, color=WHITE, size=9)
hdr_fill = PatternFill("solid", fgColor=DARK_BLUE)
bold9    = Font(name="Calibri", bold=True, size=9)
normal9  = Font(name="Calibri", size=9)


def main():
    wb = openpyxl.load_workbook(FILE)
    ws = wb["Summary"]

    # ══════════════════════════════════════════════════════════════════════════
    # CHART DATA AREA  – placed at columns O-T (hidden area below existing data)
    # We put chart data starting at row 115 to avoid overlapping existing content
    # ══════════════════════════════════════════════════════════════════════════
    chart_data_start = 115

    # ── Chart 1 data: Baseline vs Current ─────────────────────────────────────
    # Row 115: Headers
    # Row 116-120: ASM, RCA, UI IMP, HANDOVER, SAP
    ws.cell(row=chart_data_start, column=15, value="Module").font = bold9
    ws.cell(row=chart_data_start, column=16, value="Baseline").font = bold9
    ws.cell(row=chart_data_start, column=17, value="Current").font = bold9

    modules_data = [
        ("ASM", 87, 90),
        ("RCA", 37, 56),
        ("UI Improvement", 6, 6),
        ("Handover", 2, 2),
        ("SAP", 7, 8),
    ]

    for i, (mod, bl, cur) in enumerate(modules_data):
        r = chart_data_start + 1 + i
        ws.cell(row=r, column=15, value=mod).font = normal9
        ws.cell(row=r, column=16, value=bl).font = normal9
        ws.cell(row=r, column=17, value=cur).font = normal9

    # ── Chart 2 data: Change Type by Module (stacked) ─────────────────────────
    # Row 122: Headers
    # Row 123-127: Change types as rows, modules as columns
    ct_start = chart_data_start + 8  # row 123
    ws.cell(row=ct_start, column=15, value="Change Type").font = bold9
    ws.cell(row=ct_start, column=16, value="ASM").font = bold9
    ws.cell(row=ct_start, column=17, value="RCA").font = bold9
    ws.cell(row=ct_start, column=18, value="UI IMP").font = bold9
    ws.cell(row=ct_start, column=19, value="HANDOVER").font = bold9
    ws.cell(row=ct_start, column=20, value="SAP").font = bold9
    ws.cell(row=ct_start, column=21, value="Total").font = bold9

    change_data = [
        ("Added",    3, 7, 0, 0, 1),
        ("Split",    0, 13, 0, 0, 0),
        ("Moved",    1, 11, 0, 0, 1),
        ("Removed",  0, 1, 0, 0, 0),
    ]

    for i, row_data in enumerate(change_data):
        r = ct_start + 1 + i
        ct_name = row_data[0]
        vals = row_data[1:]
        ws.cell(row=r, column=15, value=ct_name).font = normal9
        for j, v in enumerate(vals):
            ws.cell(row=r, column=16 + j, value=v).font = normal9
        # Total formula
        ws.cell(row=r, column=21).value = f"=SUM(P{r}:T{r})"
        ws.cell(row=r, column=21).font = bold9

    # ══════════════════════════════════════════════════════════════════════════
    # CHART 1: Baseline vs Current (Clustered Bar)
    # ══════════════════════════════════════════════════════════════════════════
    chart1 = BarChart()
    chart1.type = "col"
    chart1.grouping = "clustered"
    chart1.title = "Baseline vs Current Scope"
    chart1.y_axis.title = "Item Count"
    chart1.x_axis.title = None
    chart1.style = 10
    chart1.width = 18
    chart1.height = 12

    cats = Reference(ws, min_col=15, min_row=chart_data_start + 1,
                     max_row=chart_data_start + 5)
    data_baseline = Reference(ws, min_col=16, min_row=chart_data_start,
                              max_row=chart_data_start + 5)
    data_current = Reference(ws, min_col=17, min_row=chart_data_start,
                             max_row=chart_data_start + 5)

    chart1.add_data(data_baseline, titles_from_data=True)
    chart1.add_data(data_current, titles_from_data=True)
    chart1.set_categories(cats)

    # Style series
    from openpyxl.chart.series import DataPoint
    from openpyxl.drawing.fill import PatternFillProperties, ColorChoice
    s1 = chart1.series[0]  # Baseline
    s1.graphicalProperties.solidFill = "4472C4"  # Blue
    s1.graphicalProperties.line.solidFill = "2F5496"
    s2 = chart1.series[1]  # Current
    s2.graphicalProperties.solidFill = "BDD7EE"  # Light blue
    s2.graphicalProperties.line.solidFill = "4472C4"

    chart1.legend.position = "b"

    # Place chart at row 5 (after executive summary, before section 2)
    ws.add_chart(chart1, "A5")

    # ══════════════════════════════════════════════════════════════════════════
    # CHART 2: Change Type by Module (Stacked Horizontal Bar)
    # ══════════════════════════════════════════════════════════════════════════
    chart2 = BarChart()
    chart2.type = "bar"  # horizontal bars
    chart2.grouping = "stacked"
    chart2.title = "Change Type by Module"
    chart2.x_axis.title = "Item Count"
    chart2.y_axis.title = None
    chart2.style = 10
    chart2.width = 18
    chart2.height = 12

    cats2 = Reference(ws, min_col=15, min_row=ct_start + 1,
                      max_row=ct_start + 4)

    # Add each change type as a series
    colors_ct = ["4472C4", "ED7D31", "FFC000", "FF0000"]  # blue, orange, gold, red
    labels_ct = ["Added", "Split", "Moved", "Removed"]

    for j, (color, label) in enumerate(zip(colors_ct, labels_ct)):
        data_ref = Reference(ws, min_col=16 + j, min_row=ct_start,
                             max_row=ct_start + 4)
        chart2.add_data(data_ref, titles_from_data=True)
        chart2.series[j].graphicalProperties.solidFill = color

    chart2.set_categories(cats2)
    chart2.legend.position = "b"
    chart2.y_axis.delete = False
    chart2.x_axis.delete = False

    # Place chart2 at row 20 (below chart 1 area)
    ws.add_chart(chart2, "A20")

    # ══════════════════════════════════════════════════════════════════════════
    # DATA TABLE next to chart 2 (row 20 area, columns H-K)
    # ══════════════════════════════════════════════════════════════════════════
    table_start_row = 36  # Place after chart 2's area
    table_start_col = 9   # Column I

    headers_t = ["Change Type", "ASM", "RCA", "UI IMP", "HANDOVER", "SAP", "Total"]
    for j, h in enumerate(headers_t):
        c = ws.cell(row=table_start_row, column=table_start_col + j, value=h)
        c.font = hdr_font
        c.fill = hdr_fill
        c.alignment = center
        c.border = thin_border

    table_data = [
        ("Added",    3, 7, 0, 0, 1),
        ("Split",    0, 13, 0, 0, 0),
        ("Moved",    1, 11, 0, 0, 1),
        ("Removed",  0, 1, 0, 0, 0),
    ]

    for i, row_data in enumerate(table_data):
        r = table_start_row + 1 + i
        ct_name = row_data[0]
        vals = row_data[1:]
        ws.cell(row=r, column=table_start_col, value=ct_name).font = bold9
        ws.cell(row=r, column=table_start_col).border = thin_border
        for j, v in enumerate(vals):
            cell = ws.cell(row=r, column=table_start_col + 1 + j, value=v)
            cell.font = normal9
            cell.alignment = center
            cell.border = thin_border
        # Total with formula
        col_start = get_column_letter(table_start_col + 1)
        col_end = get_column_letter(table_start_col + 5)
        ws.cell(row=r, column=table_start_col + 6).value = f"=SUM({col_start}{r}:{col_end}{r})"
        ws.cell(row=r, column=table_start_col + 6).font = bold9
        ws.cell(row=r, column=table_start_col + 6).alignment = center
        ws.cell(row=r, column=table_start_col + 6).border = thin_border

    # Total row
    total_r = table_start_row + 5
    ws.cell(row=total_r, column=table_start_col, value="Total").font = bold9
    ws.cell(row=total_r, column=table_start_col).fill = PatternFill("solid", fgColor=GRAY_BG)
    ws.cell(row=total_r, column=table_start_col).border = thin_border
    for j in range(1, 7):
        col_letter = get_column_letter(table_start_col + j)
        ws.cell(row=total_r, column=table_start_col + j).value = \
            f"=SUM({col_letter}{table_start_row+1}:{col_letter}{table_start_row+4})"
        ws.cell(row=total_r, column=table_start_col + j).font = bold9
        ws.cell(row=total_r, column=table_start_col + j).alignment = center
        ws.cell(row=total_r, column=table_start_col + j).border = thin_border
        ws.cell(row=total_r, column=table_start_col + j).fill = PatternFill("solid", fgColor=GRAY_BG)

    # ── Hide chart data rows (set small height) ──────────────────────────────
    for r in range(chart_data_start, ct_start + 6):
        ws.row_dimensions[r].height = 15  # compact

    # ── Save ──────────────────────────────────────────────────────────────────
    wb.save(FILE)
    print("Charts added to Summary sheet:")
    print("  Chart 1: Baseline vs Current (clustered bar) at A5")
    print("  Chart 2: Change Type by Module (stacked bar) at A20")
    print("  Data table at I36")


if __name__ == "__main__":
    main()
