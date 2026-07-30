"""Excel loading, validation, filtering and export helpers."""

from __future__ import annotations

from io import BytesIO
import re
from typing import Any, BinaryIO

import pandas as pd
from openpyxl.styles import Font, PatternFill


SHEET_NAME = "Jeneratör Verileri"

COLUMN_MAP = {
    "marka": "brand",
    "model": "model",
    "w - width (mm)": "width_mm",
    "d - depth (mm)": "depth_mm",
    "h - height (mm)": "height_mm",
    "dry ağırlık (kg)": "dry_weight_kg",
    "wet ağırlık (kg)": "wet_weight_kg",
    "güç (kw)": "power_kw",
    "faz-faz gerilim (v)": "line_voltage_v",
    "frekans (hz)": "frequency_hz",
    "yakıt tüketimi (g/bkw-h)": "fuel_consumption_g_bkwh",
    "imo": "imo",
}

REQUIRED_COLUMNS = {"brand", "model", "power_kw"}
NUMERIC_COLUMNS = [
    "width_mm",
    "depth_mm",
    "height_mm",
    "dry_weight_kg",
    "wet_weight_kg",
    "power_kw",
    "line_voltage_v",
    "frequency_hz",
    "fuel_consumption_g_bkwh",
]
RESULT_COLUMNS = [
    "brand",
    "model",
    "power_kw",
    "line_voltage_v",
    "frequency_hz",
    "width_mm",
    "depth_mm",
    "height_mm",
    "dry_weight_kg",
    "wet_weight_kg",
    "fuel_consumption_g_bkwh",
    "imo",
]
DISPLAY_NAMES = {
    "brand": "Brand",
    "model": "Model",
    "power_kw": "Power [kW]",
    "line_voltage_v": "Line-line voltage [V]",
    "frequency_hz": "Frequency [Hz]",
    "width_mm": "Width [mm]",
    "depth_mm": "Depth [mm]",
    "height_mm": "Height [mm]",
    "dry_weight_kg": "Dry weight [kg]",
    "wet_weight_kg": "Wet weight [kg]",
    "fuel_consumption_g_bkwh": "Fuel consumption [g/bkW-h]",
    "imo": "IMO",
}


class DataError(ValueError):
    """Expected workbook or input error with a readable message."""


def normalize_header(value: Any) -> str:
    """Trim, lowercase and collapse spaces while preserving underscores."""
    return re.sub(r"\s+", " ", str(value).strip().casefold())


def _read_bytes(source: bytes | bytearray | BinaryIO) -> bytes:
    if isinstance(source, (bytes, bytearray)):
        data = bytes(source)
    else:
        try:
            if hasattr(source, "seek"):
                source.seek(0)
            data = source.read()
            if hasattr(source, "seek"):
                source.seek(0)
        except Exception as exc:
            raise DataError("The uploaded Excel file could not be read.") from exc
    if not data:
        raise DataError("The uploaded Excel file is empty.")
    return data


def _parse_number(value: Any) -> float:
    if value is None or pd.isna(value) or isinstance(value, bool):
        return float("nan")
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace("\u00a0", "").replace(" ", "")
    if not text or text.casefold() in {"-", "—", "n/a", "na", "not available"}:
        return float("nan")
    match = re.match(r"^[+-]?[0-9][0-9.,]*", text)
    if not match:
        return float("nan")
    text = match.group(0)
    if "," in text and "." in text:
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif "," in text:
        text = text.replace(",", ".") if text.count(",") == 1 else text.replace(",", "")
    elif text.count(".") > 1:
        parts = text.split(".")
        text = "".join(parts[:-1]) + "." + parts[-1]
    try:
        return float(text)
    except ValueError:
        return float("nan")


def load_database(source: bytes | bytearray | BinaryIO) -> tuple[pd.DataFrame, dict[str, Any]]:
    """Read the catalogue worksheet and return usable generator rating rows."""
    try:
        excel = pd.ExcelFile(BytesIO(_read_bytes(source)), engine="openpyxl")
    except DataError:
        raise
    except Exception as exc:
        raise DataError("The uploaded file is not a readable .xlsx workbook.") from exc

    if SHEET_NAME not in excel.sheet_names:
        available = ", ".join(excel.sheet_names) or "none"
        raise DataError(
            f'The required worksheet "{SHEET_NAME}" was not found. '
            f"Available worksheets: {available}."
        )

    try:
        raw = pd.read_excel(excel, sheet_name=SHEET_NAME, dtype=object)
    except Exception as exc:
        raise DataError(f'The worksheet "{SHEET_NAME}" could not be read.') from exc

    if raw.empty and len(raw.columns) == 0:
        raise DataError(f'The worksheet "{SHEET_NAME}" does not contain a table.')

    original_rows = len(raw)
    ignored_columns = []
    kept_columns = []
    for column in raw.columns:
        normalized = normalize_header(column)
        if normalized == "index" or normalized.startswith("unnamed:"):
            ignored_columns.append(str(column))
        else:
            kept_columns.append(column)
    raw = raw.loc[:, kept_columns].copy()

    rename_map = {}
    used_names = set()
    for column in raw.columns:
        normalized = normalize_header(column)
        internal = COLUMN_MAP.get(normalized, normalized.replace(" ", "_"))
        if internal in used_names:
            raise DataError(f'Multiple Excel columns map to the same field: "{internal}".')
        used_names.add(internal)
        rename_map[column] = internal
    data = raw.rename(columns=rename_map)

    missing = REQUIRED_COLUMNS.difference(data.columns)
    if missing:
        labels = ", ".join(DISPLAY_NAMES.get(name, name) for name in sorted(missing))
        raise DataError(f"Required Excel column(s) are missing: {labels}.")

    for column in ["brand", "model", "imo"]:
        if column in data.columns:
            cleaned = data[column].astype("string").str.strip()
            data[column] = cleaned.replace({"": pd.NA, "-": pd.NA, "—": pd.NA})

    for column in NUMERIC_COLUMNS:
        if column in data.columns:
            data[column] = data[column].map(_parse_number).astype(float)

    data = data.dropna(how="all").copy()
    missing_identity = data["brand"].isna() | data["model"].isna()
    missing_power = data["power_kw"].isna()
    missing_identity_count = int(missing_identity.sum())
    invalid_power_count = int((~missing_identity & missing_power).sum())
    data = data.loc[~missing_identity & ~missing_power].copy()

    if data.empty:
        raise DataError("No usable rows contain Brand, Model, and numeric Power [kW].")

    data.insert(0, "database_row", data.index.to_series().astype(int) + 2)
    data = data.reset_index(drop=True)

    brands = sorted(data["brand"].astype(str).unique(), key=str.casefold)
    dimensions = ["width_mm", "depth_mm", "height_mm"]
    complete_dimensions = (
        int(data[dimensions].notna().all(axis=1).sum())
        if all(column in data.columns for column in dimensions)
        else 0
    )

    def present_count(column: str) -> int:
        return int(data[column].notna().sum()) if column in data.columns else 0

    stats = {
        "original_rows": original_rows,
        "usable_rows": len(data),
        "brand_count": len(brands),
        "model_count": int(data["model"].nunique()),
        "brands": brands,
        "invalid_power_rows": invalid_power_count,
        "missing_identity_rows": missing_identity_count,
        "ignored_columns": ignored_columns,
        "rows_with_voltage": present_count("line_voltage_v"),
        "rows_with_frequency": present_count("frequency_hz"),
        "rows_with_complete_dimensions": complete_dimensions,
        "rows_with_dry_weight": present_count("dry_weight_kg"),
        "rows_with_wet_weight": present_count("wet_weight_kg"),
    }
    return data, stats


def text_options(data: pd.DataFrame, column: str) -> list[str]:
    if column not in data.columns:
        return []
    values = data[column].dropna().astype(str).str.strip()
    return sorted(values[values.ne("")].unique().tolist(), key=str.casefold)


def numeric_options(data: pd.DataFrame, column: str) -> list[float]:
    if column not in data.columns:
        return []
    values = pd.to_numeric(data[column], errors="coerce").dropna().unique()
    return sorted(float(value) for value in values)


def validate_filters(filters: dict[str, Any]) -> None:
    minimum = filters.get("minimum_power_kw")
    maximum = filters.get("maximum_power_kw")
    if minimum is not None and float(minimum) <= 0:
        raise ValueError("Minimum power must be greater than zero.")
    if maximum is not None and float(maximum) <= 0:
        raise ValueError("Maximum power must be greater than zero.")
    if minimum is not None and maximum is not None and float(maximum) < float(minimum):
        raise ValueError("Maximum power cannot be lower than minimum power.")

    for key, label in {
        "maximum_width_mm": "Maximum width",
        "maximum_depth_mm": "Maximum depth",
        "maximum_height_mm": "Maximum height",
        "maximum_dry_weight_kg": "Maximum dry weight",
        "maximum_wet_weight_kg": "Maximum wet weight",
        "maximum_fuel_consumption_g_bkwh": "Maximum fuel consumption",
    }.items():
        value = filters.get(key)
        if value is not None and float(value) < 0:
            raise ValueError(f"{label} cannot be negative.")


def _numeric_match(series: pd.Series, selected: list[float]) -> pd.Series:
    numeric = pd.to_numeric(series, errors="coerce")
    result = pd.Series(False, index=series.index)
    for target in selected:
        result |= numeric.sub(float(target)).abs().le(max(1e-9, abs(float(target)) * 1e-9))
    return result


def _text_match(series: pd.Series, selected: list[str]) -> pd.Series:
    wanted = {str(value).strip().casefold() for value in selected}
    return series.astype("string").str.strip().str.casefold().isin(wanted)


def filter_catalog(
    data: pd.DataFrame, filters: dict[str, Any]
) -> tuple[pd.DataFrame, pd.DataFrame, list[str]]:
    """Return confirmed matches, unverified rows, and unavailable filter columns."""
    validate_filters(filters)
    candidate = pd.Series(True, index=data.index)
    complete = pd.Series(True, index=data.index)
    missing_columns = []

    brands = filters.get("brands") or []
    if brands:
        candidate &= _text_match(data["brand"], brands)

    query = str(filters.get("model_query") or "").strip()
    if query:
        candidate &= data["model"].astype("string").str.contains(
            re.escape(query), case=False, na=False, regex=True
        )

    minimum = filters.get("minimum_power_kw")
    maximum = filters.get("maximum_power_kw")
    if minimum is not None:
        candidate &= data["power_kw"].ge(float(minimum))
    if maximum is not None:
        candidate &= data["power_kw"].le(float(maximum))

    def optional_filter(column: str, condition: pd.Series) -> None:
        nonlocal candidate, complete
        if column not in data.columns:
            missing_columns.append(column)
            candidate &= False
            complete &= False
            return
        known = data[column].notna()
        candidate &= (~known) | condition.fillna(False)
        complete &= known

    for column, selected in {
        "line_voltage_v": filters.get("voltages_v") or [],
        "frequency_hz": filters.get("frequencies_hz") or [],
    }.items():
        if selected:
            condition = (
                _numeric_match(data[column], selected)
                if column in data.columns
                else pd.Series(False, index=data.index)
            )
            optional_filter(column, condition)

    for column, limit in {
        "width_mm": filters.get("maximum_width_mm"),
        "depth_mm": filters.get("maximum_depth_mm"),
        "height_mm": filters.get("maximum_height_mm"),
        "dry_weight_kg": filters.get("maximum_dry_weight_kg"),
        "wet_weight_kg": filters.get("maximum_wet_weight_kg"),
        "fuel_consumption_g_bkwh": filters.get("maximum_fuel_consumption_g_bkwh"),
    }.items():
        if limit is not None:
            condition = (
                data[column].le(float(limit))
                if column in data.columns
                else pd.Series(False, index=data.index)
            )
            optional_filter(column, condition)

    imo_values = filters.get("imo_values") or []
    if imo_values:
        condition = (
            _text_match(data["imo"], imo_values)
            if "imo" in data.columns
            else pd.Series(False, index=data.index)
        )
        optional_filter("imo", condition)

    exact = data.loc[candidate & complete].copy()
    unverified = data.loc[candidate & ~complete].copy()
    return _sort(exact), _sort(unverified), list(dict.fromkeys(missing_columns))


def _sort(data: pd.DataFrame) -> pd.DataFrame:
    if data.empty:
        return data.reset_index(drop=True)
    columns = [
        column
        for column in ["brand", "model", "power_kw", "frequency_hz", "line_voltage_v"]
        if column in data.columns
    ]
    return data.sort_values(columns, na_position="last", kind="mergesort").reset_index(drop=True)


def prepare_table(data: pd.DataFrame) -> pd.DataFrame:
    columns = [column for column in RESULT_COLUMNS if column in data.columns]
    result = data.loc[:, columns].copy()
    return result.rename(columns={column: DISPLAY_NAMES[column] for column in columns})


def display_missing(data: pd.DataFrame) -> pd.DataFrame:
    result = data.copy().astype(object)
    return result.where(pd.notna(result), "Not available")


def filter_summary(filters: dict[str, Any], result_count: int) -> pd.DataFrame:
    def joined(values: list[Any] | None) -> str:
        return ", ".join(str(value) for value in values) if values else "All"

    rows = [
        ("Selected brands", joined(filters.get("brands"))),
        ("Model search", filters.get("model_query") or "Not applied"),
        ("Minimum power [kW]", filters.get("minimum_power_kw")),
        ("Maximum power [kW]", filters.get("maximum_power_kw")),
        ("Selected voltages [V]", joined(filters.get("voltages_v"))),
        ("Selected frequencies [Hz]", joined(filters.get("frequencies_hz"))),
        ("Maximum width [mm]", filters.get("maximum_width_mm")),
        ("Maximum depth [mm]", filters.get("maximum_depth_mm")),
        ("Maximum height [mm]", filters.get("maximum_height_mm")),
        ("Maximum dry weight [kg]", filters.get("maximum_dry_weight_kg")),
        ("Maximum wet weight [kg]", filters.get("maximum_wet_weight_kg")),
        ("Selected IMO values", joined(filters.get("imo_values"))),
        ("Confirmed result rows", result_count),
    ]
    return pd.DataFrame(rows, columns=["Filter", "Value"])


def export_excel(
    exact: pd.DataFrame,
    filters: dict[str, Any],
    unverified: pd.DataFrame | None = None,
) -> bytes:
    """Generate the Excel export entirely in memory."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        prepare_table(exact).to_excel(writer, sheet_name="Filtered_Results", index=False)
        filter_summary(filters, len(exact)).to_excel(
            writer, sheet_name="Filter_Summary", index=False
        )
        if unverified is not None and not unverified.empty:
            prepare_table(unverified).to_excel(
                writer, sheet_name="Unverified_Missing_Data", index=False
            )

        header_fill = PatternFill("solid", fgColor="0B2545")
        header_font = Font(bold=True, color="FFFFFF")
        for sheet in writer.book.worksheets:
            sheet.freeze_panes = "A2"
            sheet.auto_filter.ref = sheet.dimensions
            for cell in sheet[1]:
                cell.fill = header_fill
                cell.font = header_font
            for cells in sheet.columns:
                longest = max(len(str(cell.value or "")) for cell in cells)
                sheet.column_dimensions[cells[0].column_letter].width = min(
                    max(longest + 2, 12), 42
                )
    return output.getvalue()
