"""Native Excel PivotTable / PivotChart / Slicer generation.

Pandas port of the reference pipeline: writes the selected columns to a raw
data workbook plus a JSON config, then drives Excel COM through
``make_pivot.ps1`` to build real PivotTables, charts, and slicers.
"""
from __future__ import annotations

import datetime as dt
import json
import math
import os
import re
import subprocess
import sys
import tempfile
import warnings
from pathlib import Path

import pandas as pd
from openpyxl import Workbook

from kpi_creator.services import parse_date_values

COM_AGG = {"mean": -4106, "sum": -4157, "count": -4112, "min": -4139, "max": -4136}
COM_CHART = {"column": 51, "line": 4, "bar": 57, "area": 1, "pie": 5}
AGG_OPTIONS = {"Average": "mean", "Sum": "sum", "Count": "count", "Min": "min", "Max": "max"}
CHART_OPTIONS = {"Column": "column", "Line": "line", "Bar": "bar", "Area": "area", "Pie": "pie"}
LAYOUT_OPTIONS = {
    "One KPI per sheet": "per_sheet",
    "Dashboard, pivots together": "all_pivot",
    "Dashboard, separate pivot sheets": "one_pivot",
}
MONTH_ABBR = ("Jan", "Feb", "Mar", "Apr", "May", "Jun",
              "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")


def resource_path(name: str) -> Path:
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parents[2]))
    return base / name


def available_dates(frame: pd.DataFrame, date_column: str) -> list[dt.date]:
    if frame is None or date_column not in frame.columns:
        return []
    parsed = parse_date_values(frame[date_column]).dropna()
    return sorted({value.date() for value in parsed})


def _safe_sheet_name(name: str, used: set[str]) -> str:
    value = re.sub(r"[\[\]:*?/\\]", "_", str(name)).strip() or "Sheet"
    value = value[:31]
    base = value
    index = 1
    while value.casefold() in used:
        suffix = "_{}".format(index)
        value = base[:31 - len(suffix)] + suffix
        index += 1
    used.add(value.casefold())
    return value


def _period_display_label(level: str, period: str) -> str:
    try:
        if level == "day":
            value = dt.date.fromisoformat(period)
            return "{:02d}-{}".format(value.day, MONTH_ABBR[value.month - 1])
        if level == "month":
            month = int(period.split("-")[-1])
            return MONTH_ABBR[month - 1]
        if level == "year":
            return "{:04d}".format(int(period))
    except (ValueError, IndexError):
        pass
    return str(period or "")


def _excel_value(value):
    if value is None:
        return None
    if isinstance(value, float) and not math.isfinite(value):
        return None
    if isinstance(value, (dict, list, tuple)):
        return json.dumps(value, ensure_ascii=False, default=str)
    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass
    if isinstance(value, pd.Timestamp):
        return value.to_pydatetime()
    if hasattr(value, "item"):
        try:
            value = value.item()
        except Exception:
            pass
    return value


def _write_raw_workbook(frame: pd.DataFrame, path: str) -> None:
    workbook = Workbook(write_only=True)
    worksheet = workbook.create_sheet("Raw Data")
    worksheet.append([str(column) for column in frame.columns])
    for row in frame.itertuples(index=False, name=None):
        worksheet.append([_excel_value(value) for value in row])
    workbook.save(path)


def _make_datetimes(frame: pd.DataFrame, date_column: str, time_column: str | None) -> pd.Series:
    dates = pd.to_datetime(parse_date_values(frame[date_column]), errors="coerce")
    if not time_column or time_column not in frame.columns:
        return dates
    values = frame[time_column]
    numeric = pd.to_numeric(values, errors="coerce")
    if numeric.notna().mean() >= 0.7:
        fraction = (numeric % 1).fillna(0.0)
        delta = pd.to_timedelta(fraction, unit="D")
    else:
        text = values.astype("string")
        parsed = pd.to_datetime(text, errors="coerce", format="ISO8601")
        if parsed.isna().all():
            with warnings.catch_warnings():
                warnings.simplefilter("ignore")
                parsed = pd.to_datetime(text, errors="coerce")
        delta = (
            pd.to_timedelta(parsed.dt.hour.fillna(0), unit="h")
            + pd.to_timedelta(parsed.dt.minute.fillna(0), unit="m")
            + pd.to_timedelta(parsed.dt.second.fillna(0), unit="s")
        )
        delta = delta.fillna(pd.Timedelta(0))
    combined = dates + delta
    return combined.where(dates.notna())


def _aggregate(series: pd.Series, function: str) -> float | None:
    values = pd.to_numeric(series, errors="coerce").dropna()
    values = values[values.abs() != float("inf")]
    if values.empty:
        return None
    if function == "mean":
        result = values.mean()
    elif function == "sum":
        result = values.sum()
    elif function == "count":
        result = len(values)
    elif function == "min":
        result = values.min()
    elif function == "max":
        result = values.max()
    else:
        raise ValueError("Unsupported aggregation: {}".format(function))
    result = float(result)
    return result if math.isfinite(result) else None


def _period_mask(date_times: pd.Series, level: str, period: str) -> pd.Series:
    if level == "day":
        target = pd.Timestamp(dt.date.fromisoformat(period))
        return date_times.dt.normalize() == target
    if level == "month":
        year, month = period.split("-", 1)
        return (date_times.dt.year == int(year)) & (date_times.dt.month == int(month))
    if level == "year":
        return date_times.dt.year == int(period)
    raise ValueError("Unknown comparison period level: {}".format(level))


def _build_summary(valid: pd.DataFrame, params: dict) -> tuple[list[str], list[dict]]:
    if not params.get("summary") or params.get("layout") == "per_sheet":
        return [], []
    period_a = params.get("period_a", "")
    period_b = params.get("period_b", "")
    if not period_a or not period_b:
        return [], []
    level = params.get("cmp_level", "month")
    date_times = valid["_DT"]
    mask_a = _period_mask(date_times, level, period_a)
    mask_b = _period_mask(date_times, level, period_b)
    group = params.get("group_col") or ""
    if group and group in valid.columns:
        groups = sorted(valid[group].dropna().unique().tolist(), key=str)
    else:
        groups = []
    if groups:
        summary_groups = [str(value) for value in groups]
        group_masks: list[pd.Series | None] = [valid[group] == value for value in groups]
    else:
        summary_groups, group_masks = ["Overall"], [None]

    summary_rows = []
    aggregation = params["agg"]
    for kpi in params["kpis"]:
        values_a, values_b, diffs, percentages = [], [], [], []
        for group_mask in group_masks:
            if group_mask is None:
                part_a, part_b = valid[mask_a], valid[mask_b]
            else:
                part_a = valid[mask_a & group_mask]
                part_b = valid[mask_b & group_mask]
            a_value = _aggregate(part_a[kpi], aggregation)
            b_value = _aggregate(part_b[kpi], aggregation)
            if a_value is None or b_value is None:
                difference = percentage = None
            else:
                difference = b_value - a_value
                percentage = None if a_value == 0 else difference / a_value * 100.0
            values_a.append(a_value)
            values_b.append(b_value)
            diffs.append(difference)
            percentages.append(percentage)
        summary_rows.append({
            "kpi": kpi, "a": values_a, "b": values_b,
            "diff": diffs, "pct": percentages,
        })
    return summary_groups, summary_rows


def run_pivot_pipeline(params: dict, progress=None) -> dict:
    def report(message: str, percent: int | None = None) -> None:
        if progress:
            progress((message, percent) if percent is not None else (message, None))

    if os.name != "nt":
        raise RuntimeError("Native PivotTable creation requires Windows and Microsoft Excel.")
    frame = params.get("frame")
    if not isinstance(frame, pd.DataFrame) or frame.empty:
        raise ValueError("The imported file has no data rows.")
    date_column = params.get("date_col")
    if not date_column or date_column not in frame.columns:
        raise ValueError("Choose a valid Date column.")
    kpis = list(params.get("kpis") or [])
    if not kpis:
        raise ValueError("Select at least one KPI value column.")
    if any(kpi not in frame.columns for kpi in kpis):
        raise ValueError("One or more selected KPI columns are not in the imported file.")

    report("Preparing data...", 5)
    date_times = _make_datetimes(frame, date_column, params.get("time_col"))
    valid = frame.copy()
    valid["_DT"] = date_times
    valid = valid[valid["_DT"].notna()]
    if valid.empty:
        raise ValueError("No valid Date/Time values were found in the selected columns.")

    group = params.get("group_col") or ""
    extra: list[str] = []
    if group:
        extra.append(group)
    for field in params.get("slicers") or []:
        if field and field not in extra:
            extra.append(field)
    output_columns = ["_DT"]
    for field in extra + kpis:
        if field not in output_columns:
            output_columns.append(field)
    pivot_data = valid[output_columns].rename(columns={"_DT": "DateTime"})
    for kpi in kpis:
        pivot_data[kpi] = pd.to_numeric(pivot_data[kpi], errors="coerce")

    report("Calculating the summary...", 25)
    summary_groups, summary_rows = _build_summary(valid, params)
    work_dir = Path(tempfile.mkdtemp(prefix="kpipivot_"))
    data_path = work_dir / "data.xlsx"
    cfg_path = work_dir / "config.json"
    script_path = work_dir / "make_pivot.ps1"
    _write_raw_workbook(pivot_data, str(data_path))

    source_path = Path(params["path"])
    output_dir = Path(params.get("out_dir") or source_path.parent)
    output_dir.mkdir(parents=True, exist_ok=True)
    stamp = dt.datetime.now().strftime("%Y-%m-%d_%H%M%S")
    output_path = output_dir / "Pivot_{}_{}.xlsx".format(source_path.stem, stamp)
    used_sheets: set[str] = set()
    config = {
        "dataPath": str(data_path.resolve()),
        "outPath": str(output_path.resolve()),
        "kpis": [{"col": kpi, "sheet": _safe_sheet_name(kpi, used_sheets)} for kpi in kpis],
        "seriesField": group,
        "slicers": list(params.get("slicers") or []),
        "aggFunc": COM_AGG[params["agg"]],
        "aggLabel": params["agg_label"],
        "chartType": COM_CHART[params["chart"]],
        "layout": params["layout"],
        "pivotSheetName": "Pivot",
        "dashboardSheetName": "Dashboard",
        "hasTime": bool(params.get("time_col")),
        "slicerItems": {
            field: int(pivot_data[field].dropna().nunique())
            for field in params.get("slicers") or [] if field in pivot_data.columns
        },
        "hasSummary": bool(summary_rows),
        "groupField": group or "Summary",
        "summaryGroups": summary_groups,
        "cmpLevel": params.get("cmp_level", "month"),
        "periodA": _period_display_label(params.get("cmp_level", "month"), params.get("period_a", "")),
        "periodB": _period_display_label(params.get("cmp_level", "month"), params.get("period_b", "")),
        "summaryRows": summary_rows,
    }
    bundled_script = resource_path("make_pivot.ps1")
    if not bundled_script.is_file():
        raise FileNotFoundError("Could not locate the Excel PivotTable automation script.")
    script_path.write_text(bundled_script.read_text(encoding="utf-8-sig"), encoding="utf-8")
    cfg_path.write_text(json.dumps(config, ensure_ascii=False, default=str), encoding="utf-8")

    powershell = Path(os.environ.get("SystemRoot", r"C:\Windows")) / "System32" / "WindowsPowerShell" / "v1.0" / "powershell.exe"
    if not powershell.is_file():
        powershell = Path("powershell.exe")
    report("Building PivotTables, charts, and slicers with Excel...", 35)
    import time

    command = [str(powershell), "-NoProfile", "-ExecutionPolicy", "Bypass",
               "-File", str(script_path), str(cfg_path)]
    process = subprocess.Popen(
        command,
        stdout=subprocess.PIPE, stderr=subprocess.PIPE,
        text=True, encoding="utf-8", errors="replace",
        creationflags=0x08000000,
    )
    start = time.monotonic()
    last_emit = 0.0
    while True:
        code = process.poll()
        if code is not None:
            break
        elapsed = time.monotonic() - start
        if elapsed > 1800:
            process.kill()
            raise RuntimeError("Excel PivotTable generation timed out.")
        percent = min(92, 35 + int(elapsed * 4.5))
        if time.monotonic() - last_emit >= 0.5:
            report("Building PivotTables, charts, and slicers with Excel...", percent)
            last_emit = time.monotonic()
        time.sleep(0.2)
    stdout, stderr = process.communicate()
    report("Finishing the workbook...", 97)
    if code != 0 or "OK" not in (stdout or ""):
        message = (stderr or stdout or "Unknown error").strip()
        raise RuntimeError("Excel PivotTable generation failed:\n" + message[-1800:])
    return {"xlsx": str(output_path)}
