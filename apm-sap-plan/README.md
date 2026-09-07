# APM SAP Plan

**SAP ECC → S/4HANA Migration Project Planning**

This project manages the APM - SAP Activity & Support workstream for the SAP Go-Live on **1 Jul 2027**.

---

## 📋 Overview

Generates professional Excel workbooks with:
- Dual-stakeholder views (SAP team + APM team)
- Weekly Gantt charts with conditional formatting
- FTE tracking grid (14 positions × 22 months)
- Auto-backup system for previous versions

---

## 🗂️ Files

| File | Purpose |
|------|---------|
| `APM SAP Plan.xlsx` | **Source plan** (read-only input) |
| `APM SAP Plan - Gantt.xlsx` | **Generated workbook** (output) |
| `build_gantt.py` | **Builder script** - generates Gantt workbook |
| `version.md` | Version log and documentation |

---

## 🔧 `build_gantt.py` Details

### What It Does

Reads `APM SAP Plan.xlsx` → Generates `APM SAP Plan - Gantt.xlsx`

### Output: Two Tabs

#### Tab 1: `SAP Timeline` (Read-Only)
- For SAP-only stakeholders
- Phase-colored SAP bars only
- Dates/status are linked formulas (auto-sync)

#### Tab 2: `APM Plan & FTE` (Editable Master)
- For APM team
- **Dual stacked bars:**
  - Phase color = SAP window
  - Blue = APM window (red TBC = not confirmed)
- **FTE grid:** Position × Month matrix
- Status dropdowns, date pickers

### Key Constants

```python
GO_LIVE = date(2027, 7, 1)      # SAP Go-Live date
GRID_START = (2026, 3)           # Grid starts March 2026
GRID_END = (2027, 12)            # Grid ends December 2027
```

### Activity Structure

Source sheet contains:
- 5 sections (A-E) grouping activities
- 12 activities with WBS numbers
- SAP dates, APM dates, mock readiness, remarks

### FTE Positions (14 total)

| Position | Count |
|----------|-------|
| ITPM | 1 |
| Scrum Master | 1 |
| BA | 1 |
| UX/UI Designer | 1 |
| Frontend Developer | 3 |
| Backend Developer | 2 |
| QA | 2 |
| QA Automation | 1 |
| Data Analytics Services | 1 |
| Data Scientist | 1 |

---

## 🕐 Timesheet Analysis

**Location:** `timesheet-analysis/`

Analyze timesheet data from `TIME_report_25_Aug.xlsx`.

### Scripts

| Script | Purpose |
|--------|---------|
| `analyze_timesheet.py` | Explore Excel structure |
| `summarize_djc.py` | Summarize Digital Job Card hours (Apr-May) |
| `summarize_djc_by_activity.py` | Group by activity type |

### Data Structure

- **Column C:** Package (filter: "Digital Job Card")
- **Column D:** Activity description
- **Column E:** Date
- **Column F:** Hours

---

## 🚀 Usage

### Regenerate Gantt Workbook

```bash
cd apm-sap-plan
python build_gantt.py
```

**Note:** Previous version is auto-backed up with timestamp.

### Run Timesheet Analysis

```bash
cd apm-sap-plan/timesheet-analysis
python analyze_timesheet.py
python summarize_djc.py
```

---

## 📊 Features

| Feature | Description |
|---------|-------------|
| Week grid | 97 weekly columns (Feb 2026 → Dec 2027) |
| Live bars | Conditional formatting (auto-updates on edit) |
| Today marker | Orange header + gray band (current week) |
| Go-Live milestone | Editable date, red block follows |
| Status dropdown | Not started / In progress / Completed |
| Date pickers | Calendar picker on all date columns |
| FTE grid | Position × Month with auto-sum total |

---

## 🤖 AI Agent Notes

**For Buffy (implementation context):**

1. **Source of truth:** `APM SAP Plan.xlsx` (read-only)
2. **Modify:** `build_gantt.py` to change output
3. **Date parsing:** `parse_dates()` handles multiple formats
4. **Conditional formatting:** Bars auto-update when dates change
5. **Regeneration:** Run `python build_gantt.py` to rebuild

---

*Last updated: 7 Sep 2026*
