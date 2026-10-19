# KPI Pivot Chart Creator

A local Windows desktop application that turns imported KPI data into **real native Excel PivotTables, PivotCharts, and slicers** — no cloud, no data upload. Everything runs on your machine, and report generation drives the Microsoft Excel that is already installed on your PC.

Author: [soemoetun](https://github.com/soemoehtun) · Tools: [smt-tool-collection](https://smt-tool-collection.vercel.app/)

---

## Contents

1. [What it does](#what-it-does)
2. [Requirements](#requirements)
3. [Install and run](#install-and-run)
4. [Build a standalone .exe](#build-a-standalone-exe)
5. [Full user guide](#full-user-guide)
6. [How the Excel output is built](#how-the-excel-output-is-built)
7. [Troubleshooting](#troubleshooting)
8. [Project layout](#project-layout)
9. [Publish a new version to GitHub](#publish-a-new-version-to-github)

---

## What it does

| Feature | Details |
| --- | --- |
| Source import | `.xlsx`, `.xlsm`, `.xlsb`, `.csv`, `.txt`; single / all / custom Excel worksheets; configure delimiter, encoding, decimal separator, null values, header row |
| Column mapping | Date column (X axis), optional Time column, Group / series field, numeric KPI value columns, slicer fields |
| Aggregation | Average, Sum, Count, Min, Max |
| Chart types | Column, Line, Bar, Area, Pie |
| Pivot layouts | One KPI per sheet · Dashboard with pivots together · Dashboard with separate pivot sheets |
| Summary table | Optional summary sheet with compare-by Day / Month / Year |
| Period comparison | Presets (previous vs latest period) plus fully custom PRE / POST period pickers |
| Progress | Live step-by-step progress bar with percentage during generation |
| Privacy | All processing stays on the machine; only Excel COM is used |

## Requirements

- **Windows 10/11** (the pipeline drives Excel through COM)
- **Microsoft Excel** (desktop) installed on the same PC
- Python 3.10 – 3.12 for running from source (a prebuilt `.exe` removes this requirement)

## Install and run

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -e .
python main.py
```

Alternative, without an editable install:

```powershell
python -m pip install -r requirements.txt
python main.py
```

Optional extras:

```powershell
python -m pip install -e ".[xlsb]"    # read .xlsb workbooks (pyxlsb)
python -m pip install -e ".[build]"   # PyInstaller for exe packaging
```

A small demo dataset is included: `sample_data/telecom_kpi_sample.csv`.

## Build a standalone .exe

Double-click **`build_exe.bat`**. It will:

1. Use (or create) the local `.venv`
2. Install app + PyInstaller dependencies
3. Produce `dist\current\KPI Pivot Chart Creator.exe`
4. Launch the exe when finished

The exe is fully standalone; end users only need **Windows + Excel** — no Python.

## Full user guide

### 1. Source file (top left)

1. Click **Browse Files** and pick a workbook/CSV.
2. For Excel files, choose a single sheet, **All Sheets**, or **Custom** (pick specific sheets).
3. For CSV/TXT, use **Import options** to set delimiter (comma, tab, pipe, custom), encoding, decimal separator (dot/comma), null values, and header row.
4. Check the **Preview rows** table to confirm the data loaded correctly.

### 2. Columns (bottom of the source card)

- **Sheet** – the worksheet to analyze.
- **Date column** – the X axis of every pivot; a Day/Month/Year summary is available when this column parses as dates.
- **Time column (optional)** – keeps time-of-day detail in the data.
- **Group / series by** – field that splits each chart into series (e.g. Region, Technology). Choose `None` for one overall table.

### 3. KPI value columns

Search and select one or more numeric metric columns (Traffic GB, Availability %, …). Each selected KPI gets its own pivot/chart sheet. Use **Select All** / **Clear** to speed up selection.

### 4. Slicer fields

Add interactive Excel slicers (e.g. Technology, Region) that filter the whole dashboard. The app prevents selecting fields that have too many distinct values to slot efficiently.

### 5. Report options

| Control | Purpose |
| --- | --- |
| **Aggregation** | How KPIs are summarized in the pivots (Average, Sum, Count, Min, Max) |
| **Chart type** | Column, Line, Bar, Area, Pie |
| **Output folder** | Where `Pivot_<filename>_<timestamp>.xlsx` is written |
| **Pivot layout** | One KPI per sheet · Dashboard, pivots together · Dashboard, separate pivot sheets |
| **Include a summary table** | Adds a summary sheet; enables the Compare panel |

### 6. Compare (summary only)

- **Compare by** – Day / Month / Year granularity.
- **Compare preset** – e.g. *Previous month vs last month*, *Day − 7*.
- **Custom** – shows PRE and POST columns with Year (`2026`), Month (`Oct`), Day (`17`) pickers listing only values present in the data.

When the dialog shows *"Pivot report created"*, the file is ready at the chosen output folder.

## How the Excel output is built

1. The selected columns are written to a temporary raw data workbook.
2. A JSON config plus a bundled PowerShell script (`make_pivot.ps1`) drives Excel COM to create genuine PivotTables, PivotCharts, slicers, and the optional summary with period comparison.
3. The workbook is saved to your output folder and the temporary data is cleaned up.

> Progress ticks to a capped 92% during the Excel stage and jumps to the end when Excel finishes — the exact speed depends on your data size and Excel.

## Troubleshooting

| Problem | Fix |
| --- | --- |
| *"Native PivotTable creation requires Windows and Microsoft Excel."* | Install desktop Excel; this app cannot run on macOS/Linux or with Excel Online. |
| *"Excel PivotTable generation failed."* | Close other Excel dialogs, then retry; the dialog shows the last Excel error lines. |
| Wrong dates / no valid dates | Re-check the Date column choice; only rows with parseable dates are used and the rest are excluded. |
| `.xlsb` fails to load | `python -m pip install pyxlsb` |

## Project layout

```text
main.py                        app entry point
build_exe.bat                  one-click exe build
make_pivot.ps1                 Excel COM automation script
pyproject.toml / requirements.txt
src/kpi_creator/
  app.py                       Qt application bootstrap
  services.py                  file import, parsing helpers
  pivot_pipeline.py            Excel pipeline + native pivot config
  ui/main_window.py            main generator window
  ui/widgets.py                multi-select panels and helpers
sample_data/telecom_kpi_sample.csv
```

---

*v1.0 — KPI Pivot Chart Creator*
