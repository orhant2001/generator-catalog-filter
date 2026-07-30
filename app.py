"""Single-page Streamlit interface for filtering a marine generator catalogue."""
 
from __future__ import annotations
 
from hashlib import sha256
from typing import Any
 
import pandas as pd
import streamlit as st
 
from data_utils import (
    DISPLAY_NAMES,
    DataError,
    export_excel,
    filter_catalog,
    load_database,
    numeric_options,
    prepare_table,
    text_options,
)
 
 
APP_TITLE = "Marine Generator Catalog Filter"
APP_DESCRIPTION = (
    "A simple catalogue filtering tool for listing marine generator product "
    "ratings from an Excel database according to user-selected technical criteria."
)
 
 
st.set_page_config(
    page_title=APP_TITLE,
    page_icon="⚙️",
    layout="wide",
    initial_sidebar_state="collapsed",
)
 
 
st.markdown(
    """
    <style>
        .stApp { background-color: #FFFFFF; color: #111827; }
        .block-container { max-width: 1500px; padding-top: 2rem; padding-bottom: 3rem; }
        h1, h2, h3 { color: #0B2545; }
        div[data-testid="stMetric"] {
            background: #F7F9FC;
            border: 1px solid #E5E7EB;
            border-radius: 10px;
            padding: 0.9rem 1rem;
        }
        div[data-testid="stFileUploader"] {
            background: #F7F9FC;
            border: 1px solid #E5E7EB;
            border-radius: 10px;
            padding: 0.8rem;
        }
        div.stButton > button[kind="primary"] {
            background-color: #1D4ED8;
            border-color: #1D4ED8;
            font-weight: 600;
        }
        div.stDownloadButton > button {
            border-color: #1D4ED8;
            color: #0B2545;
            font-weight: 600;
        }
        .catalog-note {
            background: #F7F9FC;
            border-left: 4px solid #1D4ED8;
            border-radius: 6px;
            padding: 0.85rem 1rem;
            margin: 0.5rem 0 1rem 0;
        }
        .small-muted { color: #6B7280; font-size: 0.9rem; }
    </style>
    """,
    unsafe_allow_html=True,
)
 
 
@st.cache_data(show_spinner=False)
def cached_load_database(file_bytes: bytes) -> tuple[pd.DataFrame, dict[str, Any]]:
    """Cache workbook parsing while the uploaded file remains unchanged."""
    return load_database(file_bytes)
 
 
@st.cache_data(show_spinner=False)
def cached_export(
    exact: pd.DataFrame,
    filters: dict[str, Any],
    unverified: pd.DataFrame,
) -> bytes:
    """Cache the in-memory Excel export for the current result set."""
    return export_excel(exact, filters, unverified)
 
 
def format_option(value: float) -> str:
    """Display catalogue numeric options without unnecessary trailing zeros."""
    numeric = float(value)
    return f"{numeric:,.0f}" if numeric.is_integer() else f"{numeric:,.2f}".rstrip("0").rstrip(".")
 
 
def display_table(data: pd.DataFrame) -> pd.DataFrame:
    """Prepare a readable table without altering the data used for export."""
    table = prepare_table(data).copy()
    decimal_places = {
        "Power [kW]": 1,
        "Line-line voltage [V]": 1,
        "Frequency [Hz]": 1,
        "Width [mm]": 1,
        "Depth [mm]": 1,
        "Height [mm]": 1,
        "Dry weight [kg]": 2,
        "Wet weight [kg]": 2,
        "Fuel consumption [g/bkW-h]": 2,
    }
    for column, places in decimal_places.items():
        if column in table.columns:
            table[column] = table[column].map(
                lambda value: "Not available"
                if pd.isna(value)
                else f"{float(value):,.{places}f}"
            )
    for column in table.columns:
        if column not in decimal_places:
            table[column] = table[column].astype(object).where(table[column].notna(), "Not available")
    return table
 
 
def render_grouped_results(data: pd.DataFrame) -> None:
    """Show one collapsible box per model; each box lists that model's rating rows.
 
    Rows that share a model but differ only by voltage/frequency (e.g. PP16V4000P63
    listed several times) are collapsed under a single expandable entry. When the box
    is opened, each voltage/frequency variation is shown on its own line.
    """
    grouped = data.groupby(["brand", "model"], sort=False)
    st.caption(
        f"{grouped.ngroups:,} model · {len(data):,} rating row(s) — the same model can "
        "repeat for different voltage / frequency options."
    )
 
    for (brand, model), group in grouped:
        count = len(group)
        powers = pd.to_numeric(group["power_kw"], errors="coerce").dropna()
        if powers.empty:
            power_text = ""
        elif powers.min() == powers.max():
            power_text = f" · {powers.min():,.1f} kW"
        else:
            power_text = f" · {powers.min():,.1f}–{powers.max():,.1f} kW"
 
        label = "rating" if count == 1 else "ratings"
        header = f"{brand} — {model}   ·   {count} {label}{power_text}"
 
        with st.expander(header, expanded=False):
            scenarios = display_table(group)
            # The model/brand is already in the header, so drop those columns inside.
            for redundant in ("Brand", "Model", "brand", "model"):
                if redundant in scenarios.columns:
                    scenarios = scenarios.drop(columns=redundant)
            scenarios.insert(0, "#", list(range(1, len(scenarios) + 1)))
            st.dataframe(
                scenarios,
                use_container_width=True,
                hide_index=True,
                height=min(430, 60 + 35 * min(count, 10)),
            )
 
 
FILTER_WIDGET_KEYS = (
    "all_brands",
    "selected_brands",
    "model_query",
    "selected_voltages",
    "selected_frequencies",
    "selected_imo",
    "minimum_power_enabled",
    "minimum_power_value",
    "maximum_power_enabled",
    "maximum_power_value",
    "maximum_width_enabled",
    "maximum_width_value",
    "maximum_depth_enabled",
    "maximum_depth_value",
    "maximum_height_enabled",
    "maximum_height_value",
    "maximum_dry_weight_enabled",
    "maximum_dry_weight_value",
    "maximum_wet_weight_enabled",
    "maximum_wet_weight_value",
    "maximum_fuel_enabled",
    "maximum_fuel_value",
)
 
 
def reset_results(clear_filter_widgets: bool = False) -> None:
    """Remove saved results and optionally reset filters for a new workbook."""
    for key in ("filter_result", "filter_unverified", "filter_values", "missing_filter_columns"):
        st.session_state.pop(key, None)
    if clear_filter_widgets:
        for key in FILTER_WIDGET_KEYS:
            st.session_state.pop(key, None)
 
 
def numeric_column_max(data: pd.DataFrame, column: str, fallback: float = 0.0) -> float:
    """Return a safe non-negative default maximum for an optional column."""
    if column not in data.columns:
        return fallback
    values = pd.to_numeric(data[column], errors="coerce").dropna()
    if values.empty:
        return fallback
    return max(float(values.max()), 0.0)
 
 
def optional_limit(
    checkbox_label: str,
    input_label: str,
    key_prefix: str,
    default_value: float,
    step: float,
    help_text: str | None = None,
) -> float | None:
    """Render an optional maximum/minimum numeric catalogue filter."""
    enabled = st.checkbox(checkbox_label, key=f"{key_prefix}_enabled")
    value = st.number_input(
        input_label,
        min_value=0.0,
        value=float(default_value),
        step=float(step),
        disabled=not enabled,
        key=f"{key_prefix}_value",
        help=help_text,
    )
    return float(value) if enabled else None
 
 
def render_database_summary(stats: dict[str, Any]) -> None:
    brands_text = ", ".join(stats["brands"])
    st.success(
        f'Database loaded: {stats["usable_rows"]:,} generator rating rows — '
        f'{stats["brand_count"]:,} brands.'
    )
    st.markdown(
        f'<div class="small-muted"><strong>Brands:</strong> {brands_text}</div>',
        unsafe_allow_html=True,
    )
 
    metrics = st.columns(5)
    metrics[0].metric("Rating rows", f'{stats["usable_rows"]:,}')
    metrics[1].metric("Brands", f'{stats["brand_count"]:,}')
    metrics[2].metric("Models", f'{stats["model_count"]:,}')
    metrics[3].metric("Rows with voltage", f'{stats["rows_with_voltage"]:,}')
    metrics[4].metric("Rows with frequency", f'{stats["rows_with_frequency"]:,}')
 
    with st.expander("Database coverage details", expanded=False):
        coverage = pd.DataFrame(
            [
                ("Original worksheet rows", stats["original_rows"]),
                ("Usable rating rows", stats["usable_rows"]),
                ("Rows with complete dimensions", stats["rows_with_complete_dimensions"]),
                ("Rows with dry weight", stats["rows_with_dry_weight"]),
                ("Rows with wet weight", stats["rows_with_wet_weight"]),
                ("Rows excluded for invalid/missing power", stats["invalid_power_rows"]),
                ("Rows excluded for missing brand/model", stats["missing_identity_rows"]),
            ],
            columns=["Database measure", "Rows"],
        )
        st.dataframe(coverage, use_container_width=True, hide_index=True)
        if stats["ignored_columns"]:
            st.caption("Ignored non-data columns: " + ", ".join(stats["ignored_columns"]))
 
 
def render_filter_summary(filters: dict[str, Any], exact_count: int, unverified_count: int) -> None:
    def list_text(values: list[Any] | None, all_text: str) -> str:
        return ", ".join(format_option(v) if isinstance(v, (int, float)) else str(v) for v in values) if values else all_text
 
    rows = [
        ("Brands", list_text(filters.get("brands"), "All brands")),
        ("Model search", filters.get("model_query") or "Not applied"),
        ("Minimum power [kW]", filters.get("minimum_power_kw") if filters.get("minimum_power_kw") is not None else "Not applied"),
        ("Maximum power [kW]", filters.get("maximum_power_kw") if filters.get("maximum_power_kw") is not None else "Not applied"),
        ("Voltages [V]", list_text(filters.get("voltages_v"), "All voltages")),
        ("Frequencies [Hz]", list_text(filters.get("frequencies_hz"), "All frequencies")),
        ("Confirmed matches", exact_count),
        ("Unverified due to missing filtered data", unverified_count),
    ]
    st.dataframe(pd.DataFrame(rows, columns=["Filter", "Value"]), use_container_width=True, hide_index=True)
 
 
def main() -> None:
    st.title(APP_TITLE)
    st.write(APP_DESCRIPTION)
    st.markdown(
        '<div class="catalog-note">This application filters catalogue records only. '
        "It does not size a ship electrical system, calculate generator quantity, "
        "score products, optimize selections, or provide engineering recommendations.</div>",
        unsafe_allow_html=True,
    )
 
    st.subheader("1. Upload generator database")
    uploaded_file = st.file_uploader(
        'Upload an .xlsx workbook containing the worksheet "Jeneratör Verileri".',
        type=["xlsx"],
        accept_multiple_files=False,
    )
 
    if uploaded_file is None:
        reset_results()
        st.info("Upload the prepared generator Excel database to activate the catalogue filters.")
        return
 
    file_bytes = uploaded_file.getvalue()
    file_id = sha256(file_bytes).hexdigest()
    if st.session_state.get("loaded_file_id") != file_id:
        reset_results(clear_filter_widgets=True)
        st.session_state["loaded_file_id"] = file_id
 
    try:
        with st.spinner("Reading and validating the Excel database..."):
            data, stats = cached_load_database(file_bytes)
    except DataError as exc:
        reset_results()
        st.error(str(exc))
        return
    except Exception:
        reset_results()
        st.error("The Excel database could not be processed due to an unexpected error.")
        return
 
    render_database_summary(stats)
 
    brands = text_options(data, "brand")
    voltages = numeric_options(data, "line_voltage_v")
    frequencies = numeric_options(data, "frequency_hz")
    imo_values = text_options(data, "imo")
 
    st.subheader("2. Product filters")
    top_left, top_right = st.columns(2)
    with top_left:
        all_brands = st.checkbox("All brands", value=True, key="all_brands")
        selected_brands = st.multiselect(
            "Preferred brands",
            options=brands,
            default=[],
            key="selected_brands",
            disabled=all_brands,
            help="Select one or more manufacturers, or leave All brands enabled.",
        )
        model_query = st.text_input(
            "Model search",
            placeholder="Example: M26 or K19",
            key="model_query",
            help="Case-insensitive partial model-name search.",
        )
 
    with top_right:
        selected_voltages = st.multiselect(
            "Line-line voltage [V]",
            options=voltages,
            default=[],
            key="selected_voltages",
            format_func=format_option,
            help="Leave empty to include all published voltage values.",
        )
        selected_frequencies = st.multiselect(
            "Frequency [Hz]",
            options=frequencies,
            default=[],
            key="selected_frequencies",
            format_func=format_option,
            help="Leave empty to include all published frequency values.",
        )
 
    power_left, power_right = st.columns(2)
    with power_left:
        minimum_power = optional_limit(
            "Use minimum power filter",
            "Minimum generator power [kW]",
            "minimum_power",
            default_value=float(data["power_kw"].min()),
            step=10.0,
        )
    with power_right:
        maximum_power = optional_limit(
            "Use maximum power filter",
            "Maximum generator power [kW]",
            "maximum_power",
            default_value=float(data["power_kw"].max()),
            step=10.0,
        )
 
    with st.expander("Advanced catalogue filters", expanded=False):
        st.caption(
            "When an enabled filter uses a field that is blank for a product, that product is not "
            "treated as a confirmed match. It is shown separately as unverified."
        )
        dim1, dim2, dim3 = st.columns(3)
        with dim1:
            maximum_width = optional_limit(
                "Limit maximum width",
                "Maximum width [mm]",
                "maximum_width",
                default_value=numeric_column_max(data, "width_mm"),
                step=100.0,
            )
        with dim2:
            maximum_depth = optional_limit(
                "Limit maximum depth",
                "Maximum depth [mm]",
                "maximum_depth",
                default_value=numeric_column_max(data, "depth_mm"),
                step=100.0,
            )
        with dim3:
            maximum_height = optional_limit(
                "Limit maximum height",
                "Maximum height [mm]",
                "maximum_height",
                default_value=numeric_column_max(data, "height_mm"),
                step=100.0,
            )
 
        weight1, weight2, fuel_col = st.columns(3)
        with weight1:
            maximum_dry_weight = optional_limit(
                "Limit maximum dry weight",
                "Maximum dry weight [kg]",
                "maximum_dry_weight",
                default_value=numeric_column_max(data, "dry_weight_kg"),
                step=100.0,
            )
        with weight2:
            maximum_wet_weight = optional_limit(
                "Limit maximum wet weight",
                "Maximum wet weight [kg]",
                "maximum_wet_weight",
                default_value=numeric_column_max(data, "wet_weight_kg"),
                step=100.0,
            )
        with fuel_col:
            maximum_fuel = optional_limit(
                "Limit published fuel consumption",
                "Maximum fuel consumption [g/bkW-h]",
                "maximum_fuel",
                default_value=numeric_column_max(data, "fuel_consumption_g_bkwh"),
                step=1.0,
                help_text="Uses the published catalogue value only; it is not an efficiency recommendation.",
            )
 
        selected_imo = st.multiselect(
            "IMO / emission information",
            options=imo_values,
            default=[],
            key="selected_imo",
            help="Leave empty to include all published and blank IMO values.",
        )
 
    filters = {
        "brands": [] if all_brands else selected_brands,
        "model_query": model_query.strip(),
        "minimum_power_kw": minimum_power,
        "maximum_power_kw": maximum_power,
        "voltages_v": selected_voltages,
        "frequencies_hz": selected_frequencies,
        "maximum_width_mm": maximum_width,
        "maximum_depth_mm": maximum_depth,
        "maximum_height_mm": maximum_height,
        "maximum_dry_weight_kg": maximum_dry_weight,
        "maximum_wet_weight_kg": maximum_wet_weight,
        "maximum_fuel_consumption_g_bkwh": maximum_fuel,
        "imo_values": selected_imo,
    }
 
    st.markdown("---")
    if st.button("Find generator products", type="primary", use_container_width=True):
        if not all_brands and not selected_brands:
            st.error("Select at least one brand or enable All brands.")
        else:
            try:
                exact, unverified, missing_columns = filter_catalog(data, filters)
            except ValueError as exc:
                st.error(str(exc))
            except Exception:
                st.error("The selected filters could not be applied due to an unexpected error.")
            else:
                st.session_state["filter_result"] = exact
                st.session_state["filter_unverified"] = unverified
                st.session_state["filter_values"] = filters
                st.session_state["missing_filter_columns"] = missing_columns
 
    if "filter_result" not in st.session_state:
        st.caption("Results will appear only after you click Find generator products.")
        return
 
    exact = st.session_state["filter_result"]
    unverified = st.session_state["filter_unverified"]
    applied_filters = st.session_state["filter_values"]
    missing_filter_columns = st.session_state.get("missing_filter_columns", [])
 
    st.subheader("3. Filter results")
    result_metrics = st.columns(2)
    result_metrics[0].metric("Confirmed matching rating rows", f"{len(exact):,}")
    result_metrics[1].metric("Unverified rows with missing filtered data", f"{len(unverified):,}")
 
    with st.expander("Applied filter summary", expanded=False):
        render_filter_summary(applied_filters, len(exact), len(unverified))
 
    if missing_filter_columns:
        readable = [DISPLAY_NAMES.get(column, column) for column in missing_filter_columns]
        st.warning(
            "The uploaded database does not contain the following enabled filter column(s): "
            + ", ".join(readable)
            + ". No confirmed match can be verified for those filters."
        )
 
    if exact.empty:
        st.warning(
            "No confirmed generator products were found with the selected catalogue filters. "
            "Try widening the power range, clearing one or more filters, or selecting additional brands."
        )
    else:
        render_grouped_results(exact)
 
        with st.expander("Product details and original database fields", expanded=False):
            detail_options = {
                f'{row["brand"]} — {row["model"]} — {row["power_kw"]:,.1f} kW '
                f'(Excel row {int(row["database_row"])})': index
                for index, row in exact.iterrows()
            }
            selected_label = st.selectbox("Select a rating row", options=list(detail_options))
            selected_row = exact.loc[detail_options[selected_label]].drop(labels=["database_row"], errors="ignore")
            details = pd.DataFrame(
                {
                    "Field": [DISPLAY_NAMES.get(column, column.replace("_", " ").title()) for column in selected_row.index],
                    "Value": ["Not available" if pd.isna(value) else value for value in selected_row.values],
                }
            )
            st.dataframe(details, use_container_width=True, hide_index=True)
 
    if not unverified.empty:
        with st.expander("Products with missing data for an enabled filter", expanded=False):
            st.warning(
                "These rows satisfy the known criteria, but at least one enabled filter field is blank. "
                "They are not confirmed matches and require catalogue verification."
            )
            render_grouped_results(unverified)
 
    if not exact.empty or not unverified.empty:
        try:
            export_bytes = cached_export(exact, applied_filters, unverified)
        except Exception:
            st.error("The filtered Excel export could not be generated.")
        else:
            st.download_button(
                "Export filtered results to Excel",
                data=export_bytes,
                file_name="filtered_generator_catalog_results.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
 
 
if __name__ == "__main__":
    main()
