# Citizen Property Tax & RS Assessment Survey (v25) — Live Dashboard

**Research Solutions (M&A Research Solutions LLC)** · Lahore District, Pakistan

Password-protected field-monitoring dashboard for the Citizen Property Tax &
RS (remote-sensing / satellite) Assessment Survey. Served via GitHub Pages
from `index.html`.

## What the dashboard shows

| Tab | Content |
|---|---|
| 📊 Overview | Completed vs 2,100 target, daily pace, treatment arms (T1/T2/Control), attempt disposition, E&T circle progress |
| 💻 Digital Access | Awareness & use of Punjab government portals, barriers, comfort with online forms |
| 🏛️ E&T Contact & Trust | Contact with the department, trust in officers, satisfaction, informal-payment (integrity) module |
| 📋 Records & Knowledge | Whether owners have seen their tax record, prior beliefs vs record, appeal awareness & intention, documents held |
| 🛰️ RS Technology | Comprehension checks and attitudes to satellite valuation — **compared across treatment arms** |
| 💰 Willingness to Pay | Demand curve for buying one's own RS estimate at randomised prices (Rs 0/100/200), BACE believability |
| 🗺️ Locality Tracker | Sortable, searchable, exportable completion table for all 199 sampled localities across 53 E&T circles |
| 🏠 Property & Field Ops | Property profile (floors, marla, covered area), visible assets, enumerator productivity, interview durations |

## Daily update (one click)

1. Download the fresh SurveyCTO **WIDE** export and save it in this folder as
   `Citizen Property Tax & RS Assessment Survey (v25)_WIDE.csv` (overwrite).
2. Double-click **`update_dashboard.bat`**.

The script rebuilds `index.html` from the new data and pushes it to GitHub —
the live dashboard updates within a minute.

## Files

- `build_dashboard.py` — reads the CSV export + `rs_prefill.xlsx` (sampling frame)
  + `citizen_v25.xlsx` (XLSForm, for answer labels), computes all aggregates and
  injects them into the template.
- `dashboard_template.html` — the dashboard shell (design, charts, password gate).
- `index.html` — the built dashboard (the only file GitHub Pages actually serves).
- `update_dashboard.bat` — daily one-click rebuild + push.

**Raw data files (`*.csv`, `*.xlsx`, `*.dta`) are git-ignored and never pushed** —
only the aggregated `index.html` goes to GitHub.

## Access

The dashboard is password-protected (client-side gate). The password is shared
internally by the project team — do not commit it to documentation.

---
Requires Python 3.10+ with `pandas`, `numpy`, `openpyxl`.
