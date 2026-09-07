import openpyxl
import re
from datetime import datetime
from collections import defaultdict

# Load the workbook
wb = openpyxl.load_workbook('TIME_report_25_Aug.xlsx', data_only=True)

print("Sheet names:", wb.sheetnames)

# Let's look at the first sheet to understand the structure
for sheet_name in wb.sheetnames:
    ws = wb[sheet_name]
    print(f"\n=== Sheet: {sheet_name} ===")
    print(f"Dimensions: {ws.dimensions}")
    print(f"Max row: {ws.max_row}, Max col: {ws.max_column}")
    
    # Print first 30 rows to understand structure
    for row_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=min(30, ws.max_row), values_only=False), 1):
        vals = [(cell.column_letter, cell.value) for cell in row if cell.value is not None]
        if vals:
            print(f"  Row {row_idx}: {vals}")
