import openpyxl
from datetime import datetime
from collections import defaultdict

# Load the workbook
wb = openpyxl.load_workbook('TIME_report_25_Aug.xlsx', data_only=True)
ws = wb[wb.sheetnames[0]]

# Extract all Digital Job Card entries for Apr-May
activity_summary = defaultdict(lambda: {'hours': 0, 'dates': set()})

for row in ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=False):
    package = row[2].value  # Column C = Package
    date_val = row[4].value  # Column E = Date
    hours = row[5].value  # Column F = Hours
    activity = row[3].value  # Column D = Activity description
    
    if package and 'Digital Job Card' in str(package):
        if isinstance(date_val, datetime):
            month = date_val.month
            # Filter for Apr (4) and May (5) only
            if month in [4, 5]:
                activity_name = activity if activity else 'Unknown'
                activity_summary[activity_name]['hours'] += hours if hours else 0
                activity_summary[activity_name]['dates'].add(date_val.strftime('%Y-%m-%d'))

# Sort by hours (descending)
sorted_activities = sorted(activity_summary.items(), key=lambda x: x[1]['hours'], reverse=True)

# Calculate grand total
grand_total = sum(data['hours'] for data in activity_summary.values())

# Print table
print("=" * 65)
print("Digital Job Card - Effort Breakdown by Activity (Apr-May 2026)")
print("=" * 65)
print()
print(f"{'Activity':<35} {'Hours':<10} {'%':<10} {'Days':<8}")
print("-" * 65)

for activity_name, data in sorted_activities:
    hours = data['hours']
    percentage = (hours / grand_total * 100) if grand_total > 0 else 0
    num_days = len(data['dates'])
    print(f"{activity_name:<35} {hours:<10.1f} {percentage:<10.1f} {num_days:<8}")

print("-" * 65)
print(f"{'TOTAL':<35} {grand_total:<10.1f} {'100.0':<10} ")

print()
print("=" * 65)
print("Summary:")
print(f"- Total hours worked: {grand_total:.1f} hrs")
print(f"- Total days worked: {sum(len(data['dates']) for data in activity_summary.values())}")
print(f"- Number of activity types: {len(activity_summary)}")
print("=" * 65)
