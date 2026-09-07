# APM SAP Plan - Gantt Workbook — Version Log

**Version:** v1.0.0
**Date:** 4 September 2026
**Plan basis:** Source sheet "SAP APM Plan" — *AS OF 27 AUG 2026: SAP Go Live 1 Jul 2027*

---

## 1. What this version is

`APM SAP Plan - Gantt.xlsx` — a two-tab tracking workbook for the **APM - SAP Activity & Support** workstream (APM legacy app supporting the SAP ECC → S/4HANA move), generated from the source plan by `build_gantt.py`.

### Tab 1 — `SAP Timeline` (read-only view)
- One row per activity, **phase-colored SAP bars only** (no APM rows, no FTE).
- Status, SAP dates and the Go-Live date are **linked formulas** to the master tab — this view is read-only and always in sync.
- Week columns stay narrow (width 2) — for SAP-only stakeholders.

### Tab 2 — `APM Plan & FTE` (editable master)
- **Dual stacked bars per activity:**
  - Phase color = **SAP window** (SAP Start/End, cols D:E)
  - Blue = **APM's own window** (APM Start/End, cols F:G; red `TBC` = not confirmed yet)
- **FTE Tracking grid** (below the Gantt): Position × Month, 23 monthly columns `Feb'26 → Dec'27` aligned to the Gantt month bands, 14 position rows (one row per person), auto-sum `Total FTE (all positions)` row.
- Week columns wider (~5.5) so FTE numbers display.

---

## 2. Features

| Feature | Detail |
|---|---|
| Week grid | 97 weekly columns (Mon 23 Feb 2026 → Mon 27 Dec 2027), month bands over week-start days |
| Live bars | Conditional formatting driven by the date columns — bars re-paint on edit |
| Today marker | Current-week orange header + gray band; moves with `TODAY()` |
| Go-Live milestone | Editable date row (1 Jul 2027); red block follows the date |
| Dur (d) | Live Excel formula (SAP window) |
| Status | Dropdown: Not started / In progress / Completed |
| Date input | Calendar date pickers on all four date columns (SAP + APM) |
| Mock Readiness | Dates for Mock 1–3 (May-26 / 21-Sep-26 / TBC); grey hatch = not required |
| Remark column | Free-text remarks from the source plan, wrapped |
| FTE grid | Position × Month; 1.0 = one person for the full month; Total row auto-sums |

## 3. Seeding rules & assumptions (SAP / APM dates)

- SAP Start/End parsed from source col D (`SAP Plan (approx.)`).
- APM Start/End seeded from source col F (`Involved APM Plan (approx.)`):
  - **#1 Risk, action, issue tracking** — mirrors the SAP lifecycle (Apr 26 – Sep 27).
  - **#2 CIT** — 23–28 Sep 2026 (confirmed).
  - **#7 Mock 2** — 9-Oct-26 as a single-week marker (GCP data-ready point).
  - Everything else left as red **TBC** — including SIT / UAT / Business Sim (marked *(TBC)* in source) and, per APM team decision, **Ramp Down, Blackout, Ramp Up, Hypercare** until SHANA confirms dates.
- FTE values are **MOCK SAMPLES** (from the user's FTE example, extended plausibly to Dec 2027) — to be replaced with real numbers.

## 4. Files in this repo

| File | Purpose |
|---|---|
| `APM SAP Plan.xlsx` | Source plan (read by the builder — never modified) |
| `APM SAP Plan - Gantt.xlsx` | Generated workbook (also holds manually typed FTE values) |
| `build_gantt.py` | Builder script (Python 3.12 + openpyxl) |
| `Example FTE Tracking.png`, `Example Timeline.png`, `excel architecture or layout.png` | Reference images |
| `version.md` | This log |
| `.gitignore` | Ignores `*backup*.xlsx`, `bk/`, `__pycache__/` |

## 5. Regenerating

```bash
python build_gantt.py
```
- Reads the source plan, regenerates the workbook, and **backs up the previous version automatically** (timestamped `… backup YYYYMMDD-HHMMSS.xlsx` — kept locally, not in git).
- Manual edits made inside Excel (e.g., FTE values) are overwritten on regeneration — re-apply or tell us what to bake into the script.

## 6. Change log

| Date | Change |
|---|---|
| 4 Sep 2026 | v1.0.0 — Initial versioned release. Weekly dual-bar Gantt (SAP + APM windows), live conditional-formatting bars, today/Go-Live markers, status dropdowns + date pickers, Mock Readiness + Remark columns, Position × Month FTE grid, read-only `SAP Timeline` view, auto-backups. |

---

*Maintained by the APM team. Update this log on every release/tag.*