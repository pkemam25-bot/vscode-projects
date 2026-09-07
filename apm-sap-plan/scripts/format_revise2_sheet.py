import zipfile
import shutil

SRC = "../project-info/Project Information.xlsx"
TMP = "../project-info/_tmp_revise2.xlsx"
SHEET_FILE = "xl/worksheets/sheet14.xml"

months = []
for y in (2026, 2027):
    for m in range(1, 13):
        if y == 2027 and m > 4:
            break
        months.append(f"{'Jan Feb Mar Apr May Jun Jul Aug Sep Oct Nov Dec'.split()[m-1]}'{str(y)[2:]}")

fte_vals = [0, 0.1] + [0.2] * 10 + [0.2] * 4  # Jan'26..Apr'27
assert len(months) == len(fte_vals) == 16
assert all(v <= 1 for v in fte_vals)

def col_letter(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s

def cell(ref, style, value=None, num=False):
    if value is None:
        return f'<c r="{ref}" s="{style}"/>'
    if num:
        return f'<c r="{ref}" s="{style}"><v>{value}</v></c>'
    esc = value.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    return f'<c r="{ref}" s="{style}" t="inlineStr"><is><t>{esc}</t></is></c>'

NBSP = "\u00a0"

# ---- Row 1: top border band (like original FTE sheets) ----
r1 = ['<c r="A1" s="22"/>', '<c r="B1" s="22"/>']
r1 += [cell(col_letter(n) + "1", 23, NBSP) for n in range(3, 26)]  # C..Y
rows = [f'<row r="1" spans="1:25" ht="24">{"".join(r1)}</row>']

# ---- Row 2: title ----
r2 = ['<c r="A2" s="24"/>', '<c r="B2" s="24"/>']
r2 += [cell("C2", 43, "FTE Assessment - revise2 (post go-live: conditional sign-off follow-up)")]
r2 += [cell(col_letter(n) + "2", 24) for n in range(4, 26)]  # D..Y
rows.append(f'<row r="2" spans="1:25">{"".join(r2)}</row>')

# ---- Row 3: column headers ----
r3 = ['<c r="A3" s="24"/>', '<c r="B3" s="24"/>']
r3 += [cell("C3", 45), cell("D3", 45)]                      # hidden
r3 += [cell("E3", 45, "Position"), cell("F3", 45, "Name")]  # visible
r3 += [cell("G3", 45), cell("H3", 45)]                      # hidden
r3 += [cell("I3", 168, "Vendor")]
# month headers: Jan=28, Feb=27, Mar.. = 28 (mirrors original pattern)
for i, m in enumerate(months):
    style = 27 if m == "Feb'26" else 28
    r3.append(cell(col_letter(10 + i) + "3", style, m))
rows.append(f'<row r="3" spans="1:25">{"".join(r3)}</row>')

# ---- Row 4: data (only person in row 4 of source) ----
r4 = ['<c r="A4" s="166"/>', '<c r="B4" s="46"/>']
r4 += [cell("C4", 50), cell("D4", 51)]                      # hidden
r4 += [cell("E4", 52, "ITPM"), cell("F4", 31, "Parinya Kamnoed")]
r4 += [cell("G4", 30), cell("H4", 31)]                      # hidden
r4 += [cell("I4", 169, "EPAM")]
# FTE cells: first=172, second=173, third=174, mid=175, rest=176 (mirrors original)
fte_styles = [172, 173, 174] + [175] * 5 + [176] * 8
for i, v in enumerate(fte_vals):
    r4.append(cell(col_letter(10 + i) + "4", fte_styles[i], v, num=True))
rows.append(f'<row r="4" spans="1:25">{"".join(r4)}</row>')

# ---- Row 6: note ----
rows.append('<row r="6" spans="1:25">'
            + cell("E6", 0, "Note: Project already go-live; FTE kept to follow up conditional sign-off until MA support ends (Apr 2027). Max FTE = 1.")
            + "</row>")

cols = (
    '<cols>'
    '<col min="1" max="1" width="0" hidden="1" customWidth="1"/>'
    '<col min="2" max="2" width="17" hidden="1" customWidth="1"/>'
    '<col min="3" max="3" width="45.28515625" hidden="1" customWidth="1"/>'
    '<col min="4" max="4" width="17.42578125" hidden="1" customWidth="1"/>'
    '<col min="5" max="5" width="22.42578125" customWidth="1"/>'
    '<col min="6" max="6" width="21.140625" customWidth="1"/>'
    '<col min="7" max="7" width="9.140625" hidden="1" customWidth="1"/>'
    '<col min="8" max="8" width="14.140625" hidden="1" customWidth="1"/>'
    '<col min="9" max="9" width="9.28515625" customWidth="1"/>'
    '<col min="10" max="25" width="6.14" customWidth="1"/>'
    "</cols>"
)

sheet_xml = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    '<sheetPr><tabColor rgb="FF00B050"/></sheetPr>'
    '<dimension ref="A1:Y6"/>'
    '<sheetViews><sheetView showGridLines="0" workbookViewId="0"><selection activeCell="J4" sqref="J4"/></sheetView></sheetViews>'
    '<sheetFormatPr defaultColWidth="8.85546875" defaultRowHeight="15"/>'
    + cols +
    f'<sheetData>{"".join(rows)}</sheetData>'
    "</worksheet>"
)

with zipfile.ZipFile(SRC) as zin, zipfile.ZipFile(TMP, "w", zipfile.ZIP_DEFLATED) as zout:
    for item in zin.infolist():
        if item.filename == SHEET_FILE:
            continue
        zout.writestr(item, zin.read(item.filename))
    zout.writestr(SHEET_FILE, sheet_xml)

shutil.move(TMP, SRC)
print("Done. Re-formatted sheet:", SHEET_FILE)
