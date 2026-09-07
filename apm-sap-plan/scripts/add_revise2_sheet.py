import zipfile
import shutil
import os

SRC = "../project-info/Project Information.xlsx"
TMP = "../project-info/_tmp_revise2.xlsx"

SHEET_FILE = "xl/worksheets/sheet14.xml"
RID = "rId21"
SHEET_ID = "20"
SHEET_NAME = "FTE Plan - revise2"

# Months Jan'26 .. Apr'27 (16 columns)
months = []
for y in (2026, 2027):
    for m in range(1, 13):
        if y == 2026 and m < 1:
            continue
        if y == 2027 and m > 4:
            break
        months.append(f"{'Jan Feb Mar Apr May Jun Jul Aug Sep Oct Nov Dec'.split()[m-1]}'{str(y)[2:]}")

# Row 4 of 'FTE Plan - revise': ITPM / Parinya Kamnoed / EPAM
# Jan'26=0, Feb'26=0.1, Mar-Dec'26=0.2, Jan-Apr'27=0.2 (chosen by user)
fte_vals = [0, 0.1] + [0.2] * 10 + [0.2] * 4
assert len(months) == len(fte_vals) == 16, (len(months), len(fte_vals))
assert all(v <= 1 for v in fte_vals), "FTE exceeds 1!"

def col_letter(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s

def cell(ref, value, is_text=False):
    if is_text:
        esc = value.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        return f'<c r="{ref}" t="inlineStr"><is><t>{esc}</t></is></c>'
    return f'<c r="{ref}"><v>{value}</v></c>'

# Build sheetData rows
rows_xml = []
# Header row
cells = []
cells.append(cell("A1", "Position", True))
cells.append(cell("B1", "Name", True))
cells.append(cell("C1", "Vendor", True))
for i, m in enumerate(months):
    cells.append(cell(col_letter(4 + i) + "1", m, True))
rows_xml.append(f'<row r="1" ht="15">{ "".join(cells) }</row>')

# Data row
cells = []
cells.append(cell("A2", "ITPM", True))
cells.append(cell("B2", "Parinya Kamnoed", True))
cells.append(cell("C2", "EPAM", True))
for i, v in enumerate(fte_vals):
    cells.append(cell(col_letter(4 + i) + "2", v))
rows_xml.append(f'<row r="2" ht="15">{ "".join(cells) }</row>')

# Note row
rows_xml.append(
    '<row r="4" ht="15">'
    '<c r="A4" t="inlineStr"><is><t>Note: Project already go-live; FTE kept to follow up conditional sign-off until MA support ends (Apr 2027). Max FTE = 1.</t></is></c>'
    "</row>"
)

sheet_xml = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    f'<dimension ref="A1:S4"/>'
    '<sheetViews><sheetView workbookViewId="0"/></sheetViews>'
    '<sheetFormatPr defaultRowHeight="15"/>'
    f'<sheetData>{"".join(rows_xml)}</sheetData>'
    "</worksheet>"
)

with zipfile.ZipFile(SRC) as zin, zipfile.ZipFile(TMP, "w", zipfile.ZIP_DEFLATED) as zout:
    for item in zin.infolist():
        name = item.filename
        data = zin.read(name)
        if name == "[Content_Types].xml":
            text = data.decode("utf-8")
            override = (f'<Override PartName="/{SHEET_FILE}" '
                        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>')
            assert "</Types>" in text
            text = text.replace("</Types>", override + "</Types>")
            data = text.encode("utf-8")
        elif name == "xl/_rels/workbook.xml.rels":
            text = data.decode("utf-8")
            rel = (f'<Relationship Id="{RID}" '
                   'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
                   'Target="worksheets/sheet14.xml"/>')
            assert "</Relationships>" in text
            text = text.replace("</Relationships>", rel + "</Relationships>")
            data = text.encode("utf-8")
        elif name == "xl/workbook.xml":
            text = data.decode("utf-8")
            sheet = (f'<sheet name="{SHEET_NAME}" sheetId="{SHEET_ID}" r:id="{RID}"/>')
            assert "</sheets>" in text
            text = text.replace("</sheets>", sheet + "</sheets>")
            data = text.encode("utf-8")
        elif name == SHEET_FILE:
            continue  # will be added below (shouldn't exist)
        zout.writestr(item, data)
    zout.writestr(SHEET_FILE, sheet_xml)

shutil.move(TMP, SRC)
print("Done. Added sheet:", SHEET_NAME)
