"""
Inject Excel formulas into the Summary sheet so values auto-update
when the 'baseline' sheet data changes.

Baseline data range: 'baseline'!$A$2:$A$163 (Module), $E$2:$E$163 (Sprint),
                     $I$2:$I$163 (Is Baseline), $J$2:$J$163 (Change Type)
"""

import openpyxl

FILE = r"exel file/Project Information 2.xlsx"
BL = "baseline"  # sheet name
RANGE = "$A$2:$A$163"  # row range for all columns
RANGE_E = "$E$2:$E$163"
RANGE_I = "$I$2:$I$163"
RANGE_J = "$J$2:$J$163"


def main():
    wb = openpyxl.load_workbook(FILE)
    ws = wb["Summary"]

    # ── Section 2a – Scope Overview (rows 18-27, col B = values) ──────────────
    formulas_2a = {
        # row: formula
        18: f'=COUNTIF({BL}!{RANGE_I},"baseline")',
        19: f'=COUNTA({BL}!{RANGE})',
        20: "=B19-B18",
        21: "=IF(B18=0,0,B20/B18)",
        22: f'=COUNTIF({BL}!{RANGE_J},"added")',
        23: f'=COUNTIF({BL}!{RANGE_J},"split")+COUNTIF({BL}!{RANGE_J},"Split")',
        24: f'=COUNTIF({BL}!{RANGE_J},"moved")+COUNTIF({BL}!{RANGE_J},"moved ")',
        25: f'=COUNTIF({BL}!{RANGE_J},"removed")',
        26: f'=COUNTIF({BL}!{RANGE},"sap")',
    }
    for row, formula in formulas_2a.items():
        ws.cell(row=row, column=2).value = formula
    print("Section 2a formulas added (rows 18-26)")

    # ── Section 2b – Change by Module (rows 31-36, cols B-G) ─────────────────
    # Module names in column A: ASM(31), RCA(32), UI IMPROVEMENT(33), HANDOVER(34), SAP(35)
    # We need to map display name to baseline value
    module_map = {
        31: "asm",
        32: "rca",
        33: "ui improvement",
        34: "handover",
        35: "sap",
    }
    for row, mod in module_map.items():
        m = f'"{mod}"'
        # B = Baseline items: module matches AND Is Baseline = "baseline"
        ws.cell(row=row, column=2).value = f'=COUNTIFS({BL}!{RANGE},{m},{BL}!{RANGE_I},"baseline")'
        # C = Added: module matches AND Change Type = "added"
        ws.cell(row=row, column=3).value = f'=COUNTIFS({BL}!{RANGE},{m},{BL}!{RANGE_J},"added")'
        # D = Split: module matches AND Change Type contains "split" or "Split"
        ws.cell(row=row, column=4).value = (
            f'=COUNTIFS({BL}!{RANGE},{m},{BL}!{RANGE_J},"split")'
            f'+COUNTIFS({BL}!{RANGE},{m},{BL}!{RANGE_J},"Split")'
        )
        # E = Moved: module matches AND Change Type contains "moved"
        ws.cell(row=row, column=5).value = (
            f'=COUNTIFS({BL}!{RANGE},{m},{BL}!{RANGE_J},"moved")'
            f'+COUNTIFS({BL}!{RANGE},{m},{BL}!{RANGE_J},"moved ")'
        )
        # F = Removed: module matches AND Change Type = "removed"
        ws.cell(row=row, column=6).value = f'=COUNTIFS({BL}!{RANGE},{m},{BL}!{RANGE_J},"removed")'
        # G = Current items: module matches (all items for that module)
        ws.cell(row=row, column=7).value = f'=COUNTIF({BL}!{RANGE},{m})'

    # Total row (36)
    for col in range(2, 8):
        col_letter = openpyxl.utils.get_column_letter(col)
        ws.cell(row=36, column=col).value = f"=SUM({col_letter}31:{col_letter}35)"
    print("Section 2b formulas added (rows 31-36)")

    # ── Section 6 – Gantt Sprint Total row (row 76) ──────────────────────────
    # Gantt columns: B=S18, C=S19, ... AA=S43  (cols 2-27, sprint 18-43)
    # Row 76 = Sprint Total
    sprint_start = 18
    for i in range(26):  # 26 sprints
        col = 2 + i
        sp = sprint_start + i
        ws.cell(row=76, column=col).value = f'=COUNTIF({BL}!{RANGE_E},{sp})'
    # Total of totals
    ws.cell(row=76, column=28).value = "=SUM(B76:AA76)"
    print("Section 6 Gantt total row formulas added (row 76)")

    # ── Section 7 – Sprint-by-Sprint table (rows 81-106) ─────────────────────
    # Column layout: A=Sprint label, B=Date, C=ASM, D=SAP, E=RCA, F=UI IMP, G=HANDOVER, H=Total
    sprint_modules = ["asm", "sap", "rca", "ui improvement", "handover"]
    sprint_cols = [3, 4, 5, 6, 7]  # C, D, E, F, G

    for sprint_row in range(81, 107):  # rows 81 to 106
        sp_num = sprint_row - 81 + 18  # Sprint 18 at row 81, etc.

        # C through G: COUNTIFS for each module
        for mod, col in zip(sprint_modules, sprint_cols):
            m = f'"{mod}"'
            ws.cell(row=sprint_row, column=col).value = (
                f'=COUNTIFS({BL}!{RANGE},{m},{BL}!{RANGE_E},{sp_num})'
            )

        # H = Total for sprint
        ws.cell(row=sprint_row, column=8).value = f"=SUM(C{sp_num+63}:G{sp_num+63})"

    # Grand total row (107)
    for col in range(3, 9):
        col_letter = openpyxl.utils.get_column_letter(col)
        ws.cell(row=107, column=col).value = f"=SUM({col_letter}81:{col_letter}106)"
    print("Section 7 sprint-by-sprint formulas added (rows 81-107)")

    # ── Save ──────────────────────────────────────────────────────────────────
    wb.save(FILE)
    print("\nDone. All formulas injected into Summary sheet.")


if __name__ == "__main__":
    main()
