"""
1. Backup current 'summary' sheet
2. Replace stacked chart with clustered bar chart matching the user's image:
   - Title: "Scope by Module: Baseline vs Current (items)"
   - Clustered bars: light blue (Baseline) + dark blue (Current)
   - Data labels on top of each bar
   - Clean white background
   - Legend at bottom
   - Y-axis 0-100
   - Modules: asm, rca, ui improvement, handover (no SAP)
"""

import openpyxl
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from copy import copy

FILE = r"exel file/Project Information 2_charts.xlsx"

# ── Styles ──────────────────────────────────────────────────────────────────
thin_border = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"), bottom=Side(style="thin"),
)
center = Alignment(horizontal="center", vertical="center")
bold9 = Font(name="Calibri", bold=True, size=9)
normal9 = Font(name="Calibri", size=9)
hdr_font = Font(name="Calibri", bold=True, color="FFFFFF", size=9)
hdr_fill = PatternFill("solid", fgColor="1F3864")


def main():
    wb = openpyxl.load_workbook(FILE)

    # ══════════════════════════════════════════════════════════════════════════
    # STEP 1: Backup current summary sheet
    # ══════════════════════════════════════════════════════════════════════════
    ws_src = wb["summary"]
    ws_backup = wb.copy_worksheet(ws_src)
    ws_backup.title = "summary_backup"
    print("Backup created: 'summary_backup'")

    # ══════════════════════════════════════════════════════════════════════════
    # STEP 2: Remove old chart from summary
    # ══════════════════════════════════════════════════════════════════════════
    ws = wb["summary"]
    ws._charts = []
    print("Old chart removed")

    # ══════════════════════════════════════════════════════════════════════════
    # STEP 3: New chart data at rows 10-15 (Module, Baseline, Current)
    # Using formulas so chart updates dynamically
    # ══════════════════════════════════════════════════════════════════════════
    BL = "baseline"
    RA = "$A$2:$A$163"
    RI = "$I$2:$I$163"

    # Headers at row 10
    ws.cell(row=10, column=1, value="Module").font = hdr_font
    ws.cell(row=10, column=1).fill = hdr_fill
    ws.cell(row=10, column=1).alignment = center
    ws.cell(row=10, column=1).border = thin_border

    ws.cell(row=10, column=2, value="Baseline").font = hdr_font
    ws.cell(row=10, column=2).fill = hdr_fill
    ws.cell(row=10, column=2).alignment = center
    ws.cell(row=10, column=2).border = thin_border

    ws.cell(row=10, column=3, value="Current").font = hdr_font
    ws.cell(row=10, column=3).fill = hdr_fill
    ws.cell(row=10, column=3).alignment = center
    ws.cell(row=10, column=3).border = thin_border

    # Module data with formulas
    modules = [
        ("asm", 11),
        ("rca", 12),
        ("ui improvement", 13),
        ("handover", 14),
    ]

    for mod, row in modules:
        m = f'"{mod}"'
        # Module name
        ws.cell(row=row, column=1, value=mod).font = bold9
        ws.cell(row=row, column=1).alignment = center
        ws.cell(row=row, column=1).border = thin_border

        # Baseline = COUNTIFS(module, is_baseline="baseline")
        ws.cell(row=row, column=2).value = f'=COUNTIFS({BL}!{RA},{m},{BL}!{RI},"baseline")'
        ws.cell(row=row, column=2).font = normal9
        ws.cell(row=row, column=2).alignment = center
        ws.cell(row=row, column=2).border = thin_border

        # Current = COUNTIF(module) - COUNTIF(module, removed)
        ws.cell(row=row, column=3).value = f'=COUNTIF({BL}!{RA},{m})-COUNTIFS({BL}!{RA},{m},{BL}!$J$2:$J$163,"removed")'
        ws.cell(row=row, column=3).font = normal9
        ws.cell(row=row, column=3).alignment = center
        ws.cell(row=row, column=3).border = thin_border

    print("Chart data with formulas at rows 10-14")

    # ══════════════════════════════════════════════════════════════════════════
    # STEP 4: Create clustered bar chart matching the image
    # ══════════════════════════════════════════════════════════════════════════
    chart = BarChart()
    chart.type = "col"
    chart.grouping = "clustered"
    chart.title = "Scope by Module: Baseline vs Current (items)"
    chart.y_axis.title = "Item Count"
    chart.y_axis.scaling.min = 0
    chart.y_axis.scaling.max = 100
    chart.style = 10
    chart.width = 20
    chart.height = 13

    # Categories (module names)
    cats = Reference(ws, min_col=1, min_row=11, max_row=14)

    # Baseline series (light blue)
    data_bl = Reference(ws, min_col=2, min_row=10, max_row=14)
    chart.add_data(data_bl, titles_from_data=True)
    chart.series[0].graphicalProperties.solidFill = "9DC3E6"  # Light blue
    chart.series[0].graphicalProperties.line.solidFill = "2E75B6"

    # Current series (dark blue)
    data_cur = Reference(ws, min_col=3, min_row=10, max_row=14)
    chart.add_data(data_cur, titles_from_data=True)
    chart.series[1].graphicalProperties.solidFill = "1F3864"  # Dark blue
    chart.series[1].graphicalProperties.line.solidFill = "1F3864"

    chart.set_categories(cats)

    # Data labels on top of each bar
    for s in chart.series:
        s.dLbls = DataLabelList()
        s.dLbls.showVal = True
        s.dLbls.numFmt = '0'

    # Legend at bottom
    chart.legend.position = "b"

    # Gap width between module groups
    chart.gapWidth = 100

    # Place chart at row 1 (above the sprint table)
    ws.add_chart(chart, "A16")

    print("New clustered bar chart created at A16")

    # ══════════════════════════════════════════════════════════════════════════
    # STEP 5: Save
    # ══════════════════════════════════════════════════════════════════════════
    wb.save(FILE)
    print(f"\nSaved to {FILE}")
    print("  - summary_backup: backup of original sheet")
    print("  - summary: updated with new clustered bar chart")


if __name__ == "__main__":
    main()
