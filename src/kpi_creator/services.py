from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

import pandas as pd


EXCEL_EXTENSIONS = {".xlsx", ".xlsm", ".xlsb"}
SUPPORTED_EXTENSIONS = {".csv", ".txt", *EXCEL_EXTENSIONS}


@dataclass
class ReportConfig:
    x_column: str
    time_column: str | None
    group_column: str | None
    kpis: list[str]
    slicers: list[str]
    aggregation: str
    chart_type: str
    granularity: str = "Month"
    summary_granularity: str = "Month"
    comparison_mode: str = "Previous period vs last period"
    layout: str = "Dashboard, pivots together"
    pre_period: str | None = None
    post_period: str | None = None


def workbook_sheets(path: str) -> list[str]:
    suffix = Path(path).suffix.lower()
    if suffix in {".xlsx", ".xlsm"}:
        return pd.ExcelFile(path, engine="openpyxl").sheet_names
    if suffix == ".xlsb":
        try:
            return pd.ExcelFile(path, engine="pyxlsb").sheet_names
        except ImportError as exc:
            raise RuntimeError("XLSB support needs the optional pyxlsb package. Install it with: pip install pyxlsb") from exc
    return ["Data"]


def read_dataset(path: str, sheet: str | list[str] | None = None, text_options: dict | None = None) -> pd.DataFrame:
    suffix = Path(path).suffix.lower()
    if suffix in {".csv", ".txt"}:
        options = text_options or {}
        delimiter = options.get("delimiter", "Auto-detect")
        separator = None if delimiter == "Auto-detect" else delimiter
        header = 0 if options.get("header", True) else None
        decimal = options.get("decimal", ".")
        na_values = options.get("null_values", "")
        null_list = [value.strip() for value in na_values.split(",") if value.strip()]
        try:
            frame = pd.read_csv(path, sep=separator, engine="python", encoding=options.get("encoding", "utf-8-sig"), header=header, decimal=decimal, na_values=null_list or None)
        except UnicodeDecodeError:
            frame = pd.read_csv(path, sep=separator, engine="python", encoding="latin-1", header=header, decimal=decimal, na_values=null_list or None)
        if header is None:
            frame.columns = [f"Column {index}" for index in range(1, len(frame.columns) + 1)]
    elif suffix in {".xlsx", ".xlsm"}:
        frame = _read_excel_sheets(path, sheet, "openpyxl")
    elif suffix == ".xlsb":
        try:
            frame = _read_excel_sheets(path, sheet, "pyxlsb")
        except ImportError as exc:
            raise RuntimeError("XLSB support needs the optional pyxlsb package. Install it with: pip install pyxlsb") from exc
    else:
        raise ValueError(f"Unsupported file type: {suffix or 'unknown'}")
    if not isinstance(frame, pd.DataFrame) or frame.empty and len(frame.columns) == 0:
        raise ValueError("The selected file or worksheet has no readable columns.")
    frame.columns = _unique_column_names([str(c).strip() for c in frame.columns])
    if not frame.columns.any():
        raise ValueError("The file does not contain valid column headers.")
    return frame


def _read_excel_sheets(path: str, selection: str | list[str] | None, engine: str) -> pd.DataFrame:
    if selection is None or selection == "All Sheets" or isinstance(selection, list):
        with pd.ExcelFile(path, engine=engine) as workbook:
            chosen_names = workbook.sheet_names if selection is None or selection == "All Sheets" else selection
            chosen = []
            for name in chosen_names:
                if name in workbook.sheet_names:
                    frame = pd.read_excel(workbook, sheet_name=name)
                    frame.columns = _unique_column_names([str(c).strip() for c in frame.columns])
                    chosen.append((name, frame))
        if not chosen:
            raise ValueError("Select at least one worksheet containing data.")
        schemas = [list(frame.columns) for _, frame in chosen]
        warnings = []
        if any(schema != schemas[0] for schema in schemas[1:]):
            warnings.append("Selected worksheets have different column layouts; matching header names were combined and unmatched fields were left blank.")
        source_name = "Source Worksheet"
        used_columns = {column for schema in schemas for column in schema}
        suffix = 2
        while source_name in used_columns:
            source_name = f"Source Worksheet_{suffix}"
            suffix += 1
        frames = []
        for name, frame in chosen:
            copied = frame.copy()
            copied[source_name] = name
            frames.append(copied)
        result = pd.concat(frames, ignore_index=True, sort=False)
        result.attrs["import_warnings"] = warnings
        return result
    return pd.read_excel(path, sheet_name=selection, engine=engine)


def _unique_column_names(names: list[str]) -> list[str]:
    seen: dict[str, int] = {}
    result: list[str] = []
    for index, name in enumerate(names, start=1):
        base = name or f"Column {index}"
        seen[base] = seen.get(base, 0) + 1
        result.append(base if seen[base] == 1 else f"{base}_{seen[base]}")
    return result


def numeric_columns(frame: pd.DataFrame) -> list[str]:
    result = []
    for col in frame.columns:
        converted = pd.to_numeric(frame[col], errors="coerce")
        if converted.notna().any() and converted.notna().sum() >= max(1, frame[col].notna().sum() * 0.6):
            result.append(str(col))
    return result


def suggest_date_column(frame: pd.DataFrame) -> str | None:
    candidates: list[tuple[int, str]] = []
    for col in frame.columns:
        name = str(col).lower()
        score = 2 if any(term in name for term in ("date", "day", "period", "timestamp")) else 0
        parsed = parse_date_values(frame[col].head(100))
        valid_ratio = float(parsed.notna().mean()) if len(parsed) else 0
        if valid_ratio >= 0.75:
            score += 1
        if score:
            candidates.append((score, str(col)))
    return max(candidates, default=(0, None), key=lambda item: item[0])[1]


def parse_date_values(values: pd.Series) -> pd.Series:
    """Parse common date values without interpreting ordinary numbers as timestamps."""
    if pd.api.types.is_datetime64_any_dtype(values):
        return pd.to_datetime(values, errors="coerce")
    if pd.api.types.is_numeric_dtype(values):
        # Accept the common YYYYMMDD representation, but don't treat KPI numbers
        # such as 45123 as nanoseconds since the Unix epoch.
        text = values.astype("string").str.replace(r"\.0$", "", regex=True)
        yyyymmdd = text.str.fullmatch(r"\d{8}", na=False)
        parsed = pd.Series(pd.NaT, index=values.index, dtype="datetime64[ns]")
        if yyyymmdd.any():
            parsed.loc[yyyymmdd] = pd.to_datetime(text.loc[yyyymmdd], format="%Y%m%d", errors="coerce")
        return parsed
    text = values.astype("string").str.strip()
    iso = text.str.fullmatch(r"\d{4}-\d{1,2}-\d{1,2}(?:[ T].*)?", na=False)
    parsed = pd.Series(pd.NaT, index=values.index, dtype="datetime64[ns]")
    min_ts = pd.Timestamp("1900-01-01")
    max_ts = pd.Timestamp("2100-12-31")
    if iso.any():
        iso_parsed = pd.to_datetime(text.loc[iso], errors="coerce", format="ISO8601")
        iso_parsed = iso_parsed.where((iso_parsed >= min_ts) & (iso_parsed <= max_ts))
        parsed.loc[iso] = iso_parsed
    rest = values.notna() & ~iso.fillna(False)
    if rest.any():
        rest_parsed = pd.to_datetime(text.loc[rest], errors="coerce", dayfirst=True, format="mixed")
        rest_parsed = rest_parsed.where((rest_parsed >= min_ts) & (rest_parsed <= max_ts))
        parsed.loc[rest] = rest_parsed
    return parsed


def build_summary(frame: pd.DataFrame, config: ReportConfig) -> tuple[pd.DataFrame, pd.DataFrame, list[str]]:
    if not config.kpis:
        raise ValueError("Select at least one KPI value column.")
    if config.x_column not in frame.columns:
        raise ValueError("Select a valid X-axis column.")
    data = frame.copy()
    warnings: list[str] = []
    x_values = data[config.x_column]
    x_name = config.x_column
    if config.time_column and config.time_column in data.columns:
        dates = parse_date_values(data[config.x_column])
        times = data[config.time_column].astype("string").fillna("")
        combined = pd.to_datetime(dates.dt.strftime("%Y-%m-%d") + " " + times, errors="coerce")
        if combined.notna().any():
            x_values = combined.where(combined.notna(), dates)
        else:
            warnings.append("Some date/time values could not be combined; the date column was used instead.")
    elif pd.api.types.is_datetime64_any_dtype(x_values):
        x_values = parse_date_values(x_values)
    else:
        parsed = parse_date_values(x_values)
        if parsed.notna().mean() >= 0.75:
            x_values = parsed
    data["__kpi_x"] = x_values
    invalid_axis_values = int(data["__kpi_x"].isna().sum())
    if invalid_axis_values:
        warnings.append(f"{config.x_column}: {invalid_axis_values} rows with empty or invalid X-axis values were excluded.")
    data = data[data["__kpi_x"].notna()].copy()
    if data.empty:
        raise ValueError("The selected X-axis column contains no usable values.")
    if pd.api.types.is_datetime64_any_dtype(data["__kpi_x"]):
        period = config.summary_granularity.casefold()
        if period == "day":
            data["__kpi_x"] = data["__kpi_x"].dt.normalize()
            date_format = "%d-%b-%Y"
        elif period == "month":
            data["__kpi_x"] = data["__kpi_x"].dt.to_period("M").dt.start_time
            date_format = "%b-%Y"
        elif period == "year":
            data["__kpi_x"] = data["__kpi_x"].dt.to_period("Y").dt.start_time
            date_format = "%Y"
        else:
            raise ValueError("Summary time must be Day, Month, or Year.")
    else:
        date_format = None
    dimensions = ["__kpi_x"]
    if config.group_column and config.group_column in data.columns and config.group_column not in config.kpis:
        dimensions.append(config.group_column)
    for kpi in config.kpis:
        missing = int(data[kpi].isna().sum())
        if missing:
            warnings.append(f"{kpi}: {missing} source values are blank and excluded from numeric aggregation.")
        numeric = pd.to_numeric(data[kpi], errors="coerce")
        invalid = data[kpi].notna() & numeric.isna()
        if invalid.any():
            warnings.append(f"{kpi}: {int(invalid.sum())} non-numeric values excluded from aggregation.")
        data[kpi] = numeric
    agg = {"average": "mean", "sum": "sum", "count": "count", "min": "min", "max": "max"}.get(config.aggregation.casefold())
    if agg is None:
        raise ValueError(f"Unsupported aggregation: {config.aggregation}")
    # Use one reducer for the selected values so KPI aggregation follows the
    # same DataFrameGroupBy path regardless of the source column names.
    summary = data.groupby(dimensions, dropna=False, sort=True)[config.kpis].agg(agg).reset_index()
    summary = summary.rename(columns={"__kpi_x": x_name})
    if pd.api.types.is_datetime64_any_dtype(summary[x_name]):
        summary = summary.sort_values(x_name, kind="stable")
        summary[x_name] = summary[x_name].dt.strftime(date_format or "%d-%b-%Y")
    else:
        summary = summary.sort_values(x_name, kind="stable")
    pivot = summary.copy()
    return summary.reset_index(drop=True), pivot.reset_index(drop=True), warnings


def apply_slicer_filters(frame: pd.DataFrame, values: dict[str, list[str]]) -> pd.DataFrame:
    filtered = frame
    for column, selected in values.items():
        if selected and column in filtered.columns:
            as_text = filtered[column].astype("string").fillna("(blank)")
            filtered = filtered[as_text.isin(selected)]
    return filtered


def build_comparisons(frame: pd.DataFrame, config: ReportConfig) -> dict[str, pd.DataFrame]:
    if config.x_column not in frame.columns or not config.kpis:
        return {"Daily": pd.DataFrame(), "Monthly": pd.DataFrame(), "Yearly": pd.DataFrame()}
    data = frame.copy()
    data["__date"] = parse_date_values(data[config.x_column])
    data = data[data["__date"].notna()]
    if data.empty:
        return {"Daily": pd.DataFrame(), "Monthly": pd.DataFrame(), "Yearly": pd.DataFrame()}
    group_col = config.group_column if config.group_column in data.columns else None
    result: dict[str, pd.DataFrame] = {}
    for period in ("Daily", "Monthly", "Yearly"):
        if period == "Daily":
            data["__period"] = data["__date"].dt.normalize()
            offset = pd.DateOffset(days=1)
        elif period == "Monthly":
            data["__period"] = data["__date"].dt.to_period("M").dt.start_time
            offset = pd.DateOffset(months=1)
        else:
            data["__period"] = data["__date"].dt.to_period("Y").dt.start_time
            offset = pd.DateOffset(years=1)
        for kpi in config.kpis:
            data[kpi] = pd.to_numeric(data[kpi], errors="coerce")
        latest = data["__period"].max()
        post_period = latest
        pre_period = latest - offset
        selected_period = {"Day": "Daily", "Month": "Monthly", "Year": "Yearly"}.get(config.granularity)
        if period == selected_period:
            mode = config.comparison_mode.casefold()
            if "custom" in mode and config.pre_period and config.post_period:
                pre_period, post_period = _custom_periods(config.pre_period, config.post_period, period)
            elif period == "Daily" and "7" in mode:
                pre_period = latest - pd.DateOffset(days=7)
        # Keep calendar-period comparisons exact; don't substitute a prior observed period.
        group_keys = ["__period"] + ([group_col] if group_col else [])
        agg_method = config.aggregation.lower()
        agg_method = {"average": "mean", "count": "count"}.get(agg_method, agg_method)
        pre = data[data["__period"] == pre_period].groupby(group_keys, dropna=False)[config.kpis].agg(agg_method)
        post = data[data["__period"] == post_period].groupby(group_keys, dropna=False)[config.kpis].agg(agg_method)
        if group_col:
            try:
                aligned = pre.join(post, lsuffix="_pre", rsuffix="_post", how="outer").reset_index()
            except Exception:
                aligned = pd.DataFrame()
        else:
            pre_values = pre.iloc[0] if not pre.empty else pd.Series(dtype=float)
            post_values = post.iloc[0] if not post.empty else pd.Series(dtype=float)
            aligned = pd.DataFrame([{}])
            for kpi in config.kpis:
                aligned[f"{kpi}_pre"] = [pre_values.get(kpi, float("nan"))]
                aligned[f"{kpi}_post"] = [post_values.get(kpi, float("nan"))]
        rows: list[dict] = []
        for _, record in aligned.iterrows():
            for kpi in config.kpis:
                before = record.get(f"{kpi}_pre")
                after = record.get(f"{kpi}_post")
                try:
                    before_num = float(before)
                    after_num = float(after)
                    missing = pd.isna(before_num) or pd.isna(after_num)
                except (TypeError, ValueError):
                    missing = True
                    before_num = after_num = float("nan")
                delta = after_num - before_num if not missing else float("nan")
                pct = "N/A" if not missing and before_num == 0 else ((delta / before_num) * 100 if not missing else float("nan"))
                row = {"KPI": kpi}
                if group_col:
                    row["Group"] = record.get(group_col)
                row.update({
                    "PRE Period": _format_period(pre_period, period),
                    "POST Period": _format_period(post_period, period),
                    "PRE Value": before_num,
                    "POST Value": after_num,
                    "Absolute Change": delta,
                    "Percentage Change": pct,
                    "Status": "Available" if not missing else "Unavailable",
                })
                rows.append(row)
        result[period] = pd.DataFrame(rows)
    return result


def _format_period(value: pd.Timestamp, kind: str) -> str:
    if kind == "Daily":
        return value.strftime("%d-%b-%Y")
    if kind == "Monthly":
        return value.strftime("%b-%Y")
    return value.strftime("%Y")


def _custom_periods(pre: str, post: str, kind: str) -> tuple[pd.Timestamp, pd.Timestamp]:
    pre_date = pd.Timestamp(pre)
    post_date = pd.Timestamp(post)
    if kind == "Daily":
        return pre_date.normalize(), post_date.normalize()
    if kind == "Monthly":
        return pre_date.to_period("M").start_time, post_date.to_period("M").start_time
    return pre_date.to_period("Y").start_time, post_date.to_period("Y").start_time


def export_excel(
    path: str,
    source_name: str,
    summary: pd.DataFrame,
    comparisons: dict[str, pd.DataFrame],
    quality: pd.DataFrame,
    layout: str = "Dashboard, pivots together",
    kpis: Iterable[str] | None = None,
) -> None:
    with pd.ExcelWriter(path, engine="xlsxwriter") as writer:
        workbook = writer.book
        header = workbook.add_format({"bold": True, "font_color": "#FFFFFF", "bg_color": "#0F766E", "border": 0, "align": "left"})
        title = workbook.add_format({"bold": True, "font_size": 16, "font_color": "#0F766E"})
        used_names: set[str] = set()
        if layout == "Dashboard, pivots together" or not kpis:
            _write_formatted(writer, _excel_sheet_name("Pivot_Summary", used_names), summary, header, title, f"KPI Pivot Summary · {source_name}")
        else:
            for kpi in kpis:
                if kpi not in summary.columns:
                    continue
                dimensions = [col for col in summary.columns if col not in set(kpis)]
                sheet = _excel_sheet_name(kpi, used_names)
                table = summary[dimensions + [kpi]]
                _write_formatted(writer, sheet, table, header, title, f"{kpi} Summary · {source_name}")
        for period, table in comparisons.items():
            if not table.empty:
                _write_formatted(writer, _excel_sheet_name(f"{period}_Comparison", used_names), table, header, title, f"{period} KPI Comparison")
        if not quality.empty:
            _write_formatted(writer, _excel_sheet_name("Data_Quality", used_names), quality, header, title, "Data Quality Notes")


def _excel_sheet_name(value: str, used: set[str]) -> str:
    safe = "".join("_" if char in "[]:*?/\\" else char for char in str(value)).strip("' ") or "Summary"
    safe = safe[:31]
    candidate = safe
    suffix = 2
    while candidate.casefold() in {name.casefold() for name in used}:
        tail = f"_{suffix}"
        candidate = f"{safe[:31 - len(tail)]}{tail}"
        suffix += 1
    used.add(candidate)
    return candidate


def _write_formatted(writer, sheet_name: str, table: pd.DataFrame, header, title, heading: str) -> None:
    table.to_excel(writer, sheet_name=sheet_name, startrow=2, index=False)
    ws = writer.sheets[sheet_name]
    ws.write(0, 0, heading, title)
    ws.freeze_panes(3, 0)
    if len(table.columns):
        ws.autofilter(2, 0, max(2, len(table) + 2), len(table.columns) - 1)
        for col_index, col in enumerate(table.columns):
            width = max(12, min(34, max([len(str(col)), *(len(str(v)) for v in table[col].head(100))]) + 2))
            ws.set_column(col_index, col_index, width)
            ws.write(2, col_index, str(col), header)
