"""
Create a 'Summary' sheet in 'Project Information 2.xlsx'
- Section 1: Executive Summary for PMO (timeline status, RCA, SAP)
- Section 2: Dashboard – Baseline vs Current by Module + Change Type breakdown
- Section 3: SAP Activity Highlights
Insert the sheet before the existing 'dashboard' sheet.
"""

import openpyxl
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, numbers
)
from openpyxl.utils import get_column_letter
from collections import Counter, defaultdict
from copy import copy

FILE = r"exel file/Project Information 2.xlsx"

# ── Colours / Styles ────────────────────────────────────────────────────────
DARK_BLUE  = "1F3864"
MED_BLUE   = "2E75B6"
LIGHT_BLUE = "D6E4F0"
GREEN_BG   = "C6EFCE"
GREEN_FT   = "006100"
YELLOW_BG  = "FFEB9C"
YELLOW_FT  = "9C6500"
ORANGE_BG  = "FCD5B4"
RED_BG     = "FFC7CE"
RED_FT     = "9C0006"
WHITE      = "FFFFFF"
GRAY_BG    = "F2F2F2"
SAP_BG     = "FFF2CC"   # light yellow for SAP highlight
SAP_FONT   = "BF8F00"   # dark gold for SAP

hdr_font   = Font(name="Calibri", bold=True, color=WHITE, size=11)
hdr_fill   = PatternFill("solid", fgColor=DARK_BLUE)
sub_font   = Font(name="Calibri", bold=True, color=DARK_BLUE, size=10)
sub_fill   = PatternFill("solid", fgColor=LIGHT_BLUE)
normal     = Font(name="Calibri", size=10)
bold10     = Font(name="Calibri", bold=True, size=10)
title_font = Font(name="Calibri", bold=True, color=DARK_BLUE, size=14)
section_font = Font(name="Calibri", bold=True, color=WHITE, size=11)
section_fill = PatternFill("solid", fgColor=MED_BLUE)
green_fill = PatternFill("solid", fgColor=GREEN_BG)
green_font = Font(name="Calibri", bold=True, color=GREEN_FT, size=11)
yellow_fill = PatternFill("solid", fgColor=YELLOW_BG)
yellow_font = Font(name="Calibri", bold=True, color=YELLOW_FT, size=11)
sap_fill   = PatternFill("solid", fgColor=SAP_BG)
sap_font   = Font(name="Calibri", bold=True, color=SAP_FONT, size=10)
thin_border = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"), bottom=Side(style="thin"),
)
center = Alignment(horizontal="center", vertical="center", wrap_text=True)
left_wrap = Alignment(horizontal="left", vertical="center", wrap_text=True)


def apply_border(ws, row, cols):
    for c in range(1, cols + 1):
        ws.cell(row=row, column=c).border = thin_border


def section_header(ws, row, text, cols=7):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=cols)
    cell = ws.cell(row=row, column=1, value=text)
    cell.font = section_font
    cell.fill = section_fill
    cell.alignment = Alignment(horizontal="left", vertical="center")
    for c in range(1, cols + 1):
        ws.cell(row=row, column=c).fill = section_fill
        ws.cell(row=row, column=c).border = thin_border


def table_header(ws, row, headers, start_col=1):
    for i, h in enumerate(headers):
        c = ws.cell(row=row, column=start_col + i, value=h)
        c.font = hdr_font
        c.fill = hdr_fill
        c.alignment = center
        c.border = thin_border


def write_row(ws, row, values, start_col=1, font=normal, alignment=left_wrap, fills=None):
    for i, v in enumerate(values):
        c = ws.cell(row=row, column=start_col + i, value=v)
        c.font = font
        c.alignment = alignment
        c.border = thin_border
        if fills and i < len(fills) and fills[i]:
            c.fill = fills[i]


def main():
    wb = openpyxl.load_workbook(FILE)

    # ── Read baseline data ────────────────────────────────────────────────────
    ws_bl = wb["baseline"]
    rows_data = []
    for row in ws_bl.iter_rows(min_row=2, max_row=ws_bl.max_row, min_col=1, max_col=11, values_only=True):
        if row[0] is not None or row[2] is not None:
            rows_data.append({
                "module": row[0],
                "feature": row[1],
                "story_id": row[2],
                "description": row[3],
                "sprint": row[4],
                "start": row[5],
                "end": row[6],
                "status": row[7],
                "is_baseline": row[8],
                "change_type": row[9],
                "note": row[10],
            })

    baseline_items = [r for r in rows_data if r["is_baseline"] == "baseline"]
    total_items = len(rows_data)
    baseline_count = len(baseline_items)

    baseline_modules = Counter(r["module"] for r in baseline_items)
    all_modules = Counter(r["module"] for r in rows_data)
    # SAP is highlighted separately, so we exclude it from the "main" module counts for dashboard
    # but include it in the total
    sap_items = [r for r in rows_data if r["module"] == "sap"]
    sap_count = len(sap_items)

    # Change type aggregation (case-insensitive)
    change_map = {"added": "added", "moved": "moved", "removed": "removed", "split": "split"}
    change_types_raw = Counter()
    for r in rows_data:
        ct = r["change_type"]
        if ct:
            change_types_raw[change_map.get(ct.strip().lower(), ct.strip().lower())] += 1
    # Note: "moved " (trailing space) normalises to "moved"

    added_count = change_types_raw.get("added", 0)
    moved_count = change_types_raw.get("moved", 0)
    removed_count = change_types_raw.get("removed", 0)
    split_count = change_types_raw.get("split", 0)

    # Per-module change breakdown
    module_list = ["asm", "rca", "ui improvement", "handover", "sap"]

    mod_change = defaultdict(lambda: {"baseline": 0, "current": 0, "added": 0, "removed": 0, "moved": 0, "split": 0})
    for r in rows_data:
        m = r["module"]
        mod_change[m]["current"] += 1
        if r["is_baseline"] == "baseline":
            mod_change[m]["baseline"] += 1
        ct = r["change_type"]
        if ct:
            key = change_map.get(ct.strip().lower(), ct.strip().lower())
            mod_change[m][key] += 1

    # Sprint items for RCA section
    sprint_range = f"{min(r['sprint'] for r in rows_data if r['sprint'])} – {max(r['sprint'] for r in rows_data if r['sprint'])}"

    # ── Create Summary sheet ──────────────────────────────────────────────────
    ws = wb.create_sheet("Summary", wb.sheetnames.index("dashboard"))

    # Column widths
    col_widths = {1: 38, 2: 18, 3: 18, 4: 18, 5: 22, 6: 18, 7: 18, 8: 20}
    for col, w in col_widths.items():
        ws.column_dimensions[get_column_letter(col)].width = w

    row = 1

    # ── TITLE ─────────────────────────────────────────────────────────────────
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=7)
    c = ws.cell(row=row, column=1, value="APM Project – Summary & Dashboard")
    c.font = title_font
    c.alignment = Alignment(horizontal="left", vertical="center")
    row += 1

    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=7)
    c = ws.cell(row=row, column=1, value="Project: S1 AI Advisory Flowline Back Pressure Optimization  |  ITPM: Parinya Kamnoed (Chaii)")
    c.font = Font(name="Calibri", size=10, italic=True, color="555555")
    c.alignment = Alignment(horizontal="left")
    row += 1

    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=7)
    c = ws.cell(row=row, column=1, value="Baseline established: April 2026  |  As of: 24 Aug 2026")
    c.font = Font(name="Calibri", size=10, italic=True, color="555555")
    c.alignment = Alignment(horizontal="left")
    row += 2

    # ══════════════════════════════════════════════════════════════════════════
    # SECTION 1  –  EXECUTIVE SUMMARY  (for PMO)
    # ══════════════════════════════════════════════════════════════════════════
    section_header(ws, row, "1  |  EXECUTIVE SUMMARY – Project Status (PMO View)", 7)
    row += 1

    # Status badge
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=7)
    status_cell = ws.cell(row=row, column=1,
                          value="[OK]  PROJECT ON TRACK  –  Timeline unchanged, RCA on plan to finish Sprint 35, SAP activities accommodated without impact")
    status_cell.font = green_font
    status_cell.fill = green_fill
    status_cell.alignment = Alignment(horizontal="center", vertical="center")
    for c_idx in range(1, 8):
        ws.cell(row=row, column=c_idx).fill = green_fill
        ws.cell(row=row, column=c_idx).border = thin_border
    ws.row_dimensions[row].height = 28
    row += 2

    # Key messages table
    headers = ["Item", "Detail", "Impact on Timeline"]
    table_header(ws, row, headers)
    row += 1

    messages = [
        ("Project Timeline (Sprint 18 – 43)",
         "05 Jan 2026 – 25 Dec 2026",
         "NOT affected – unchanged from baseline"),
        ("RCA Module Completion",
         "RCA still planned to complete at Sprint 35",
         "NOT affected – on schedule"),
        ("SAP Activities (ECC → S/4HANA)",
         f"{sap_count} SAP items across Sprint 20–30",
         "NOT affected – accommodated within existing timeline"),
        ("Change Types Impact",
         f"Added: {added_count} | Moved: {moved_count} | Removed: {removed_count} | Split: {split_count}",
         "NOT affected – all within original timeline"),
        ("Customer/User Focus",
         "User focus from RCA onward (Sprint 24+); ASM already passed & accepted",
         "N/A – ASM phase complete"),
    ]

    for msg in messages:
        fills = [None, None, green_fill]
        write_row(ws, row, msg, font=bold10, fills=fills)
        # Make the "Impact" column green
        ws.cell(row=row, column=3).font = Font(name="Calibri", bold=True, color=GREEN_FT, size=10)
        row += 1

    row += 1

    # ══════════════════════════════════════════════════════════════════════════
    # SECTION 2  –  DASHBOARD: Baseline vs Current by Module
    # ══════════════════════════════════════════════════════════════════════════
    section_header(ws, row, "2  |  DASHBOARD – Baseline vs Current Scope", 7)
    row += 1

    # 2a – Scope summary
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=7)
    ws.cell(row=row, column=1, value="2a  |  Scope Overview").font = sub_font
    ws.cell(row=row, column=1).fill = sub_fill
    for c_idx in range(1, 8):
        ws.cell(row=row, column=c_idx).fill = sub_fill
        ws.cell(row=row, column=c_idx).border = thin_border
    row += 1

    headers = ["Metric", "Value", "Comment"]
    table_header(ws, row, headers)
    row += 1

    scope_rows = [
        ("Baseline scope (items)", baseline_count, "Original baseline at SOR 2026"),
        ("Current scope (items)", total_items, "Everything in the plan today"),
        ("Net change (items)", total_items - baseline_count, "Track of change"),
        ("Net change (%)", f"{(total_items - baseline_count) / baseline_count * 100:.1f}%", ""),
        ("   - New add-on requests (Added)", added_count, "New scope items"),
        ("   - Items broken down (Split)", split_count, "Same scope, more detailed tickets"),
        ("   - Items moved to another sprint", moved_count, "Sequence changed, scope unchanged"),
        ("   - Items removed (de-scoped)", removed_count, "Removed from plan"),
        ("SAP Activities (highlighted)", sap_count, "ECC → S/4HANA migration tasks"),
        ("Timeline (start - end)", "05 Jan 26 - 25 Dec 26", "Unchanged from baseline"),
    ]
    for m in scope_rows:
        write_row(ws, row, m)
        row += 1

    row += 1

    # 2b – Change by Module
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=7)
    ws.cell(row=row, column=1, value="2b  |  Change Breakdown by Module").font = sub_font
    ws.cell(row=row, column=1).fill = sub_fill
    for c_idx in range(1, 8):
        ws.cell(row=row, column=c_idx).fill = sub_fill
        ws.cell(row=row, column=c_idx).border = thin_border
    row += 1

    headers = ["Module", "Baseline Items", "Added", "Split", "Moved", "Removed", "Current Items"]
    table_header(ws, row, headers)
    row += 1

    total_baseline = 0
    total_current = 0
    total_added = 0
    total_split = 0
    total_moved = 0
    total_removed = 0

    for mod in module_list:
        d = mod_change[mod]
        vals = [
            mod.upper(),
            d["baseline"],
            d["added"],
            d["split"],
            d["moved"],
            d["removed"],
            d["current"],
        ]
        fills = [sap_fill if mod == "sap" else None] + [None] * 6
        fonts = [sap_font if mod == "sap" else bold10] + [normal] * 6
        write_row(ws, row, vals, font=normal, fills=fills)
        if mod == "sap":
            ws.cell(row=row, column=1).font = sap_font
        row += 1
        total_baseline += d["baseline"]
        total_current += d["current"]
        total_added += d["added"]
        total_split += d["split"]
        total_moved += d["moved"]
        total_removed += d["removed"]

    # Total row
    total_vals = ["TOTAL", total_baseline, total_added, total_split, total_moved, total_removed, total_current]
    write_row(ws, row, total_vals, font=bold10,
              fills=[PatternFill("solid", fgColor=GRAY_BG)] * 7)
    row += 2

    # ══════════════════════════════════════════════════════════════════════════
    # SECTION 3  –  SAP ACTIVITY HIGHLIGHTS
    # ══════════════════════════════════════════════════════════════════════════
    section_header(ws, row, "3  |  SAP ACTIVITY HIGHLIGHTS  (ECC → S/4HANA Migration)", 7)
    row += 1

    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=7)
    note_cell = ws.cell(row=row, column=1,
        value="Note: SAP items are highlighted in the 'baseline' sheet (module = 'sap'). "
              "These activities relate to ECC to S/4HANA migration and have been accommodated within the existing sprint plan without timeline impact.")
    note_cell.font = Font(name="Calibri", size=9, italic=True, color=SAP_FONT)
    note_cell.alignment = left_wrap
    ws.row_dimensions[row].height = 30
    row += 1

    headers = ["Sprint", "Story ID", "Feature", "Description", "Status", "Change Type", "Note"]
    table_header(ws, row, headers)
    row += 1

    for s in sap_items:
        vals = [
            f"Sprint {s['sprint']}",
            s["story_id"],
            s["feature"],
            s["description"],
            "Done" if s["status"] == 1 else "In Progress",
            s["change_type"] if s["change_type"] else "baseline",
            s["note"] if s["note"] else "",
        ]
        fills = [sap_fill] * 7
        write_row(ws, row, vals, font=normal, fills=fills)
        row += 1

    row += 1

    # ══════════════════════════════════════════════════════════════════════════
    # SECTION 4  –  EXCLUSIVE SUMMARY  (PMO Message)
    # ══════════════════════════════════════════════════════════════════════════
    section_header(ws, row, "4  |  EXCLUSIVE SUMMARY – PMO Assurance Message", 7)
    row += 1

    ws.merge_cells(start_row=row, start_column=1, end_row=row + 6, end_column=7)
    summary_text = (
        "To PMO & Stakeholders,\n\n"
        "The APM project (S1 AI Advisory - Flowline Back Pressure Optimization) "
        "remains ON TRACK per the original baseline established in April 2026.\n\n"
        "Key points:\n"
        "  1.  Project timeline (Sprint 18 - 43, ending 25 Dec 2026) is UNCHANGED.\n"
        "  2.  RCA module is on plan to complete at Sprint 35 as originally scheduled.\n"
        "  3.  SAP activities (ECC to S/4HANA migration) have been integrated into "
        "Sprints 20-30 and do NOT impact the project timeline.\n"
        "  4.  All change types (Added, Moved, Removed, Split) have been absorbed "
        "within the original plan without extending the schedule.\n"
        "  5.  ASM module (Sprint 18-27) is COMPLETE. User/customer focus is now "
        "on RCA and subsequent modules.\n\n"
        "The project remains green - no timeline extension required."
    )
    c = ws.cell(row=row, column=1, value=summary_text)
    c.font = Font(name="Calibri", size=10)
    c.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    c.fill = PatternFill("solid", fgColor=GREEN_BG)
    for r_idx in range(row, row + 7):
        for c_idx in range(1, 8):
            ws.cell(row=r_idx, column=c_idx).fill = PatternFill("solid", fgColor=GREEN_BG)
            ws.cell(row=r_idx, column=c_idx).border = thin_border
    ws.row_dimensions[row].height = 180
    row += 8

    # ══════════════════════════════════════════════════════════════════════════
    # SECTION 5  –  LEGEND / NOTES
    # ══════════════════════════════════════════════════════════════════════════
    section_header(ws, row, "5  |  REFERENCE & NOTES", 7)
    row += 1

    notes = [
        ("Data Source", "'baseline' sheet – item-level plan with Is Baseline and Change Type columns"),
        ("Sprint Calendar", "'sprintday' sheet – Sprint 18 (05 Jan 2026) to Sprint 43 (25 Dec 2026)"),
        ("Change Types", "Added = new scope, Moved = sprint re-sequence, Removed = de-scoped, Split = broken down into sub-tickets"),
        ("SAP Module", "SAP tasks relate to ECC to S/4HANA migration; highlighted in yellow on the 'baseline' sheet"),
        ("RCA Focus", "Customer/user focus starts from RCA (Sprint 24 onward); ASM (Sprint 18-27) already passed and accepted"),
    ]

    headers = ["Item", "Detail"]
    table_header(ws, row, headers)
    row += 1

    for n in notes:
        write_row(ws, row, n)
        row += 1

    # ── Freeze panes & print setup ────────────────────────────────────────────
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A4"

    # ── Save ──────────────────────────────────────────────────────────────────
    wb.save(FILE)
    print(f"OK  Summary sheet created and inserted before 'dashboard' in '{FILE}'")
    print(f"   Total rows in baseline: {total_items}")
    print(f"   Baseline items: {baseline_count}")
    print(f"   SAP items: {sap_count}")
    print(f"   Change types – Added: {added_count}, Moved: {moved_count}, Removed: {removed_count}, Split: {split_count}")


if __name__ == "__main__":
    main()
