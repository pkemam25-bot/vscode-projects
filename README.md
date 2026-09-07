# VS Code Projects Repository

This repository contains multiple project folders organized under the `~/Desktop/VS Code` directory. Each folder represents a separate project with its own purpose and codebase.

---

## 📁 Project Structure

```
VS Code/
├── apm-sap-plan/          # SAP project planning & Gantt charts
├── excel improvement/     # Excel file manipulation scripts
└── timesheet analysis/    # Timesheet data analysis tools
```

---

## 🎯 Project 1: APM SAP Plan

**Location:** `apm-sap-plan/`  
**Purpose:** SAP ECC → S/4HANA migration project planning with Gantt charts, FTE tracking, and timesheet analysis.

### Overview
This project manages the **APM - SAP Activity & Support** workstream for the SAP Go-Live on **1 Jul 2027**. It generates professional Excel workbooks with dual-stakeholder views (SAP team + APM team) and includes timesheet analysis for resource tracking.

### Key Files

| File | Purpose |
|------|---------|
| `APM SAP Plan.xlsx` | **Source plan** (read-only input) - contains raw activity data |
| `APM SAP Plan - Gantt.xlsx` | **Generated workbook** (output) - dual-tab Gantt + FTE tracking |
| `build_gantt.py` | **Builder script** - generates the Gantt workbook from source |
| `version.md` | Version log and documentation |

### `build_gantt.py` - What It Does

This Python script (Python 3.12 + openpyxl) generates a two-tab Excel workbook:

#### Tab 1: `SAP Timeline` (Read-Only View)
- For **SAP-only stakeholders**
- One row per activity with **phase-colored SAP bars only**
- Status + dates are **linked formulas** to the master tab (always in sync)
- Narrow week columns (width 2) for compact view

#### Tab 2: `APM Plan & FTE` (Editable Master)
- For the **APM team**
- **Dual stacked bars per activity:**
  - Phase color = SAP window (SAP Start/End)
  - Blue = APM's own window (APM Start/End; red TBC = not confirmed)
- **FTE Tracking grid:** Position × Month matrix (Feb'26 → Dec'27)
- Status dropdowns, date pickers, conditional formatting
- Go-Live milestone marker (1 Jul 2027)

### Activity Structure

The source sheet (`SAP APM Plan`) contains:
- **Section headers** (A-E) grouping related activities
- **12 activities** with WBS numbers, status, SAP dates, APM dates, mock readiness, and remarks

### FTE Positions Tracked

| Position | Description |
|----------|-------------|
| ITPM | IT Project Manager |
| Scrum Master | Agile facilitator |
| BA | Business Analyst |
| UX/UI Designer | User experience |
| Frontend Developer (3) | Frontend team |
| Backend Developer (2) | Backend team |
| QA (2) | Quality assurance |
| QA Automation | Test automation |
| Data Analytics Services | Data support |
| Data Scientist | ML/AI support |

### How to Regenerate

```bash
cd apm-sap-plan
python build_gantt.py
```

- Reads `APM SAP Plan.xlsx` (source)
- Generates `APM SAP Plan - Gantt.xlsx` (output)
- Auto-backs up previous version with timestamp

### `timesheet-analysis/` Sub-Folder

**Location:** `apm-sap-plan/timesheet-analysis/`  
**Purpose:** Analyze timesheet data from `TIME_report_25_Aug.xlsx`

#### Key Scripts

| Script | Purpose |
|--------|---------|
| `analyze_timesheet.py` | Explores and displays the structure of the timesheet Excel file |
| `summarize_djc.py` | Summarizes "Digital Job Card" entries for Apr-May 2026 |
| `summarize_djc_by_activity.py` | Groups DJC entries by activity type |

#### Data Structure
- **Column C:** Package name (filters for "Digital Job Card")
- **Column D:** Activity description
- **Column E:** Date (datetime)
- **Column F:** Hours worked

#### What `summarize_djc.py` Does
1. Loads `TIME_report_25_Aug.xlsx`
2. Filters rows where Package = "Digital Job Card"
3. Extracts Apr-May 2026 entries
4. Groups by month and date
5. Prints daily and monthly hour totals

---

## 📊 Project 2: Excel Improvement

**Location:** `excel improvement/`  
**Purpose:** General-purpose Excel file manipulation scripts for improving Excel workbooks.

### Overview
A collection of Python scripts for:
- Adding charts to Excel files
- Adding formulas
- Creating Gantt charts
- Revising/formatting worksheets
- Reading and analyzing Excel structures

### Key Scripts

| Script | Purpose |
|--------|---------|
| `add_charts.py` | Add charts to Excel workbooks |
| `add_formulas.py` | Insert formulas into worksheets |
| `add_gantt.py` | Create Gantt chart visualizations |
| `add_revise2_sheet.py` | Add a revised version sheet |
| `build_gantt.py` | Build Gantt charts from data |
| `create_summary.py` | Create summary sheets |
| `fix_formulas.py` | Fix broken formulas |
| `fix_gantt_dates.py` | Correct Gantt date calculations |
| `format_revise2_sheet.py` | Format the revised sheet |
| `gantt.py` | Gantt chart generation |
| `read_xlsx.py` | Read and display Excel structure |
| `revise_chart.py` | Revise existing charts |
| `update_dates_and_fte.py` | Update date and FTE values |

### `exel file/` Sub-Folder

Contains the actual Excel files being manipulated:
- `Project Information.xlsx` - Main project file
- `Project Information 2.xlsx` - Second version
- `Project Information 2_final.xlsx` - Final version
- `Project Information - original.xlsx` - Original backup
- `bk/` - Backup folder

---

## ⏰ Project 3: Timesheet Analysis

**Location:** `timesheet analysis/`  
**Purpose:** Standalone timesheet analysis (similar to the sub-folder in apm-sap-plan).

### Key Files

| File | Purpose |
|------|---------|
| `TIME_report_25_Aug.xlsx` | Timesheet data source |
| `analyze_timesheet.py` | Explore and display Excel structure |
| `summarize_djc.py` | Summarize Digital Job Card hours |
| `summarize_djc_by_activity.py` | Group by activity type |

### Usage

```bash
cd "timesheet analysis"
python analyze_timesheet.py      # View structure
python summarize_djc.py          # Summarize DJC hours
python summarize_djc_by_activity.py  # Group by activity
```

---

## 🔧 Technical Details

### Dependencies

All projects use:
- **Python 3.12+**
- **openpyxl** - Excel file manipulation

Install with:
```bash
pip install openpyxl
```

### File Conventions

- Excel files: `.xlsx` format
- Python scripts: snake_case naming
- Generated files include version timestamps in backups
- Source files are never modified (read-only input)

---

## 🤖 AI Agent Context

**For Buffy (AI Agent) - Implementation Notes:**

### When Working on APM SAP Plan:
1. **Source of truth:** `APM SAP Plan.xlsx` - read only
2. **Generated output:** `APM SAP Plan - Gantt.xlsx` - can be regenerated
3. **Builder script:** `build_gantt.py` - modify this to change output
4. **Key constants:**
   - `GO_LIVE = date(2027, 7, 1)` - SAP Go-Live date
   - `GRID_START = (2026, 3)` - Grid starts March 2026
   - `GRID_END = (2027, 12)` - Grid ends December 2027
   - `POSITIONS` - List of FTE positions with sample values
5. **Date parsing:** `parse_dates()` handles multiple formats (e.g., "23-28 Sep 2026", "Apr 26 - Sep 2027")
6. **Conditional formatting:** Bars auto-update when dates change

### When Working on Excel Improvement:
1. Scripts are modular - each handles one task
2. `read_xlsx.py` is useful for exploring new Excel files
3. `gantt.py` contains a standalone Gantt renderer (terminal output)

### When Working on Timesheet Analysis:
1. Data source: `TIME_report_25_Aug.xlsx`
2. Key filter: Package column = "Digital Job Card"
3. Date range: Apr-May 2026 (configurable)
4. Output: Console summary (can be extended to Excel)

---

## 📝 Version History

| Date | Project | Change |
|------|---------|--------|
| 4 Sep 2026 | apm-sap-plan | v1.0.0 - Initial Gantt workbook with dual tabs |
| 7 Sep 2026 | All | Initial repository commit with all projects |

---

## 👥 Maintainer

- **Parinya Kamnoed** (parinya_kamnoed@epam.com)
- APM Team

---

*Generated with Codebuff 🤖*
