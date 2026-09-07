import sys
import zipfile
import xml.etree.ElementTree as ET

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

NS = {
    "m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}

def col_letter(ref):
    """Convert a cell reference like 'B12' to column letter 'B'."""
    letters = []
    for ch in ref:
        if ch.isalpha():
            letters.append(ch)
        else:
            break
    return "".join(letters)

def read_workbook(zf):
    root = ET.fromstring(zf.read("xl/workbook.xml"))
    sheets = []
    for sh in root.find("m:sheets", NS):
        sheets.append((sh.get("name"), sh.get("sheetId"), sh.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")))
    return sheets

def sheet_paths(zf):
    """Map relationship id -> worksheet path from workbook.xml.rels."""
    rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    mapping = {}
    for rel in rels:
        if rel.get("Type", "").endswith("/worksheet"):
            mapping[rel.get("Id")] = rel.get("Target")
    return mapping

def shared_strings(zf):
    if "xl/sharedStrings.xml" not in zf.namelist():
        return None
    root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    strings = []
    for si in root.findall("m:si", NS):
        # Concatenate all text runs (handles rich text <r><t>...</t></r>)
        text = "".join(t.text or "" for t in si.iter("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t"))
        strings.append(text)
    return strings

def read_sheet(zf, path, strings):
    if not path.startswith("xl/"):
        path = "xl/" + path.lstrip("/")
    root = ET.fromstring(zf.read(path))
    sheet_data = root.find("m:sheetData", NS)
    rows = []
    for row in sheet_data.findall("m:row", NS):
        cells = {}
        for c in row.findall("m:c", NS):
            ref = c.get("r")
            t = c.get("t")
            v = c.find("m:v", NS)
            isel = c.find("m:is", NS)
            if t == "s" and v is not None and strings is not None:
                val = strings[int(v.text)]
            elif t == "inlineStr" and isel is not None:
                val = "".join(x.text or "" for x in isel.iter("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t"))
            elif v is not None:
                val = v.text
            else:
                val = ""
            cells[ref] = val
        rows.append(cells)
    return rows

def main(path):
    with zipfile.ZipFile(path) as zf:
        sheets = read_workbook(zf)
        paths = sheet_paths(zf)
        strings = shared_strings(zf)
        print(f"Workbook: {path}")
        print(f"Sheets ({len(sheets)}): " + ", ".join(s[0] for s in sheets))
        print("=" * 60)
        for name, sheet_id, rid in sheets:
            p = paths.get(rid)
            if p is None:
                print(f"\n--- Sheet '{name}': (no worksheet found) ---")
                continue
            print(f"\n--- Sheet '{name}' ---")
            rows = read_sheet(zf, p, strings)
            for row in rows:
                if not row:
                    continue
                # Sort by column letter then row number, skip empty cells
                items = sorted(row.items(), key=lambda kv: (col_letter(kv[0]), int("".join(ch for ch in kv[0] if ch.isdigit()) or 0)))
                nonempty = [(ref, val) for ref, val in items if str(val).strip()]
                if not nonempty:
                    continue
                print(" | ".join(f"{ref}={val}" for ref, val in nonempty))

if __name__ == "__main__":
    main(sys.argv[1])
