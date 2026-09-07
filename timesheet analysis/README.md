# Timesheet Analysis

**Timesheet data analysis and summarization**

Analyze timesheet reports, extract Digital Job Card entries, and summarize hours by month and activity.

---

## 📋 Overview

Processes timesheet Excel files to:
- Explore and understand Excel structure
- Filter for "Digital Job Card" entries
- Summarize hours by month and date
- Group by activity type

---

## 🗂️ Files

| File | Purpose |
|------|---------|
| `TIME_report_25_Aug.xlsx` | Timesheet data source |
| `analyze_timesheet.py` | Explore Excel structure |
| `summarize_djc.py` | Summarize DJC hours (Apr-May) |
| `summarize_djc_by_activity.py` | Group by activity type |

---

## 📊 Data Structure

The timesheet Excel file has these key columns:

| Column | Letter | Description |
|--------|--------|-------------|
| Package | C | Filter target: "Digital Job Card" |
| Activity | D | Activity description |
| Date | E | Date of work (datetime) |
| Hours | F | Hours worked |

---

## 🚀 Usage

### Explore Structure

```bash
cd "timesheet analysis"
python analyze_timesheet.py
```

Shows: Sheet names, dimensions, first 30 rows with values.

### Summarize Digital Job Card Hours

```bash
python summarize_djc.py
```

**Output:**
```
Digital Job Card - Time Summary (Apr-May)
==================================================

Month        Date            Time (hrs)  
----------------------------------------
April        2026-04-02      8.0         
April        2026-04-03      7.5         
...
April Total: 120.0 hrs

May          2026-05-01      8.0         
...
May Total: 95.0 hrs

Grand Total: 215.0 hrs
```

### Group by Activity

```bash
python summarize_djc_by_activity.py
```

Shows hours grouped by activity type.

---

## 🔧 Customization

### Change Date Range

Edit `summarize_djc.py`:

```python
# Filter for Apr (4) and May (5) only
if month in [4, 5]:
```

Change to other months:
```python
# Filter for Jun-Sep
if month in [6, 7, 8, 9]:
```

### Change Package Filter

Edit the filter condition:

```python
if 'Digital Job Card' in str(package):
```

Change to:
```python
if 'Other Package' in str(package):
```

---

## 🤖 AI Agent Notes

**For Buffy (implementation context):**

1. **Data source:** `TIME_report_25_Aug.xlsx`
2. **Key filter:** Package column = "Digital Job Card"
3. **Date range:** Configurable in script (currently Apr-May)
4. **Output:** Console summary (can extend to Excel)
5. **Dependencies:** Only openpyxl

---

*Last updated: 7 Sep 2026*
