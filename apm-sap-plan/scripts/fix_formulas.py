import openpyxl

# Load the workbook
wb = openpyxl.load_workbook('../project-info/Project Information 2_final.xlsx')

# The external file path prefix to find and replace (match both old and new paths)
external_prefix = "'C:\\Users\\ParinyaKamnoed\\Desktop\\VS Code\\excel improvement\\exel file\\[Project Information 2_final.xlsx]baseline'!"

fixed_count = 0

# Check all sheets in the workbook
for sheet_name in wb.sheetnames:
    ws = wb[sheet_name]
    for row in ws.iter_rows():
        for cell in row:
            if cell.value and isinstance(cell.value, str) and cell.value.startswith('=') and external_prefix in cell.value:
                old_val = cell.value
                # Replace the external path with just 'baseline'!
                cell.value = cell.value.replace(external_prefix, "'baseline'!")
                fixed_count += 1
                print(f"  Fixed {sheet_name}!{cell.coordinate}: {old_val[:80]}...")

# Save
wb.save('../project-info/Project Information 2_final.xlsx')
print(f"\nTotal fixed: {fixed_count} formulas across all sheets")
