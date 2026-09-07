# Excel Improvement

**General-purpose Excel file manipulation scripts**

A collection of Python scripts for improving Excel workbooks with charts, formulas, Gantt visualizations, and formatting.

---

## 📋 Overview

Modular scripts that can be used independently or combined:
- Add charts to Excel files
- Insert formulas
- Create Gantt charts
- Revise and format worksheets
- Read and analyze Excel structures

---

## 🗂️ Scripts

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
| `gantt.py` | Gantt chart generation (terminal) |
| `read_xlsx.py` | Read and display Excel structure |
| `revise_chart.py` | Revise existing charts |
| `update_dates_and_fte.py` | Update date and FTE values |

---

## 📁 Excel Files

**Location:** `exel file/`

| File | Description |
|------|-------------|
| `Project Information.xlsx` | Main project file |
| `Project Information 2.xlsx` | Second version |
| `Project Information 2_final.xlsx` | Final version |
| `Project Information - original.xlsx` | Original backup |
| `bk/` | Backup folder |

---

## 🚀 Usage

### Read Excel Structure

```bash
cd "excel improvement"
python read_xlsx.py
```

### Generate Gantt (Terminal)

```bash
python gantt.py
```

Output: ASCII Gantt chart in terminal.

### Add Charts

```bash
python add_charts.py
```

### Fix Formulas

```bash
python fix_formulas.py
```

---

## 🔧 Technical Details

### Dependencies

- Python 3.12+
- openpyxl

### File Naming Convention

- Scripts: `snake_case.py`
- Excel files: `Title Case.xlsx`
- Backups: `bk/` folder

---

## 🤖 AI Agent Notes

**For Buffy (implementation context):**

1. **Modular design:** Each script handles one task
2. **Start with:** `read_xlsx.py` to explore new Excel files
3. **Gantt:** `gantt.py` is standalone (terminal output)
4. **Formulas:** Use `add_formulas.py` as reference for formula patterns
5. **Backups:** Check `bk/` folder for previous versions

---

*Last updated: 7 Sep 2026*
