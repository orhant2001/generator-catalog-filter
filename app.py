"""Single-page Streamlit interface for filtering a marine generator catalogue."""
 
from __future__ import annotations
 
import re
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
    "Filter marine generator product ratings from an Excel catalogue by technical criteria."
)
SCOPE_NOTE = (
    "Catalogue filtering only — it lists published product ratings and does not size a ship "
    "electrical system, choose quantities, or provide engineering recommendations."
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
        .block-container { max-width: 1400px; padding-top: 2rem; padding-bottom: 3rem; }
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
        .small-muted { color: #6B7280; font-size: 0.9rem; }
 
        /* Section labels for the ##### headers used inside the filter form */
        h5 {
            color: #1D4ED8;
            font-size: 0.82rem !important;
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 0.06em;
            border-bottom: 1px solid #E5E7EB;
            padding-bottom: 0.35rem;
            margin: 1.2rem 0 0.6rem 0;
        }
 
        /* Rounded, bordered expanders */
        div[data-testid="stExpander"] {
            border: 1px solid #E5E7EB;
            border-radius: 12px;
            background: #FFFFFF;
            overflow: hidden;
        }
        div[data-testid="stExpander"] summary { font-weight: 600; color: #0B2545; }
 
        /* Blue chips for multiselect selections */
        span[data-baseweb="tag"] {
            background-color: #1D4ED8 !important;
            border-radius: 6px !important;
        }
 
        /* Blue slider handle */
        div[data-testid="stSlider"] [data-baseweb="slider"] div[role="slider"] {
            background-color: #1D4ED8 !important;
            border-color: #1D4ED8 !important;
        }
 
        div[data-testid="stNumberInput"] input { border-radius: 8px; }
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
 
 
_IMO_TIER_RE = re.compile(r"\b(III|II|I)\b")
_IMO_TIER_ORDER = ("I", "II", "III")
 
 
def imo_tiers_in(value: Any) -> list[str]:
    """Return the IMO NOx tiers (I, II, III) mentioned in a raw catalogue value.
 
    Longest-first alternation with word boundaries means "IMO Tier III" yields
    ["III"] (not ["II", "I"]) and "IMO Tier II / IMO Tier III" yields ["II", "III"].
    """
    return _IMO_TIER_RE.findall(str(value).upper())
 
 
def build_imo_options(data: pd.DataFrame) -> tuple[list[str], dict[str, set[str]]]:
    """Build tier-based IMO filter options mapped to the raw values they include.
 
    Selecting "IMO Tier II" also includes combined values such as
    "IMO Tier II / IMO Tier III" — without altering the stored or displayed data,
    because each option expands to the exact raw strings it should match.
    Values with no recognizable tier are kept as their own literal option.
    """
    raw_values = text_options(data, "imo")
    tier_to_raw: dict[str, set[str]] = {}
    literal_values: list[str] = []
    for raw in raw_values:
        tiers = set(imo_tiers_in(raw))
        if tiers:
            for tier in tiers:
                tier_to_raw.setdefault(tier, set()).add(raw)
        else:
            literal_values.append(raw)
 
    options: list[str] = []
    option_to_raw: dict[str, set[str]] = {}
 
    ordered_tiers = [tier for tier in _IMO_TIER_ORDER if tier in tier_to_raw]
    ordered_tiers += [tier for tier in sorted(tier_to_raw) if tier not in _IMO_TIER_ORDER]
    for tier in ordered_tiers:
        label = f"IMO Tier {tier}"
        options.append(label)
        option_to_raw[label] = tier_to_raw[tier]
 
    for raw in literal_values:
        options.append(raw)
        option_to_raw[raw] = {raw}
 
    return options, option_to_raw
 
 
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
    st.caption(f"{grouped.ngroups:,} model · {len(data):,} rating row(s)")
 
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
    "selected_brands",
    "model_query",
    "selected_voltages",
    "selected_frequencies",
    "power_range",
    "selected_imo",
    "max_width_mm",
    "max_depth_mm",
    "max_height_mm",
    "max_dry_weight_kg",
    "max_wet_weight_kg",
    "max_fuel_consumption_g_bkwh",
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
 
 
def max_number_filter(
    label: str,
    column: str,
    data: pd.DataFrame,
    step: float,
    help_text: str | None = None,
) -> float | None:
    """Render an optional maximum filter as a single blank-able number input.
 
    Leaving the field empty means the filter is not applied — no checkbox needed.
    """
    col_max = numeric_column_max(data, column)
    hint = "Leave blank for no limit."
    if col_max:
        hint += f" Catalogue maximum: {col_max:,.0f}."
    if help_text:
        hint = f"{help_text} {hint}"
    value = st.number_input(
        label,
        min_value=0.0,
        value=None,
        step=float(step),
        key=f"max_{column}",
        help=hint,
        placeholder="No limit",
    )
    return float(value) if value is not None else None
 
 
def render_filters(data: pd.DataFrame, imo_options: list[str]) -> dict[str, Any]:
    """Render the full filter form and return a raw (pre-expansion) selections dict."""
    brands = text_options(data, "brand")
    voltages = numeric_options(data, "line_voltage_v")
    frequencies = numeric_options(data, "frequency_hz")
 
    st.markdown("##### Identification")
    id_left, id_right = st.columns(2)
    with id_left:
        selected_brands = st.multiselect(
            "Brands",
            options=brands,
            default=[],
            key="selected_brands",
            placeholder="All brands — leave empty for every manufacturer",
            help="Pick one or more manufacturers, or leave empty to include all brands.",
        )
    with id_right:
        model_query = st.text_input(
            "Model search",
            placeholder="e.g. M26, K19, PP16V4000",
            key="model_query",
            help="Case-insensitive partial model-name search.",
        )
 
    st.markdown("##### Electrical")
    elec_left, elec_right = st.columns(2)
    with elec_left:
        selected_voltages = st.multiselect(
            "Line-line voltage [V]",
            options=voltages,
            default=[],
            key="selected_voltages",
            format_func=format_option,
            placeholder="All voltages",
            help="Leave empty to include all published voltage values.",
        )
    with elec_right:
        selected_frequencies = st.multiselect(
            "Frequency [Hz]",
            options=frequencies,
            default=[],
            key="selected_frequencies",
            format_func=format_option,
            placeholder="All frequencies",
            help="Leave empty to include all published frequency values.",
        )
 
    st.markdown("##### Power")
    power_min = float(data["power_kw"].min())
    power_max = float(data["power_kw"].max())
    if power_min >= power_max:
        st.info(f"Every rating row shares the same power: {power_min:,.1f} kW.")
        minimum_power = None
        maximum_power = None
    else:
        power_low, power_high = st.slider(
            "Power range [kW]",
            min_value=power_min,
            max_value=power_max,
            value=(power_min, power_max),
            step=max(round((power_max - power_min) / 100, 1), 0.1),
            key="power_range",
            help="Drag the handles to set a minimum and/or maximum. Full range means no power filter.",
        )
        minimum_power = power_low if power_low > power_min else None
        maximum_power = power_high if power_high < power_max else None
 
    st.markdown("##### Size, weight, fuel & emissions")
    st.caption(
        "Enter a maximum value to apply a limit, or leave a field blank to ignore it. "
        "If an applied limit uses a field that is blank for a product, that product is listed "
        "separately as unverified rather than as a confirmed match."
    )
 
    st.markdown("**Maximum dimensions [mm]**")
    dim1, dim2, dim3 = st.columns(3)
    with dim1:
        maximum_width = max_number_filter("Max width [mm]", "width_mm", data, step=50.0)
    with dim2:
        maximum_depth = max_number_filter("Max depth [mm]", "depth_mm", data, step=50.0)
    with dim3:
        maximum_height = max_number_filter("Max height [mm]", "height_mm", data, step=50.0)
 
    st.markdown("**Maximum weight [kg]**")
    weight1, weight2 = st.columns(2)
    with weight1:
        maximum_dry_weight = max_number_filter("Max dry weight [kg]", "dry_weight_kg", data, step=50.0)
    with weight2:
        maximum_wet_weight = max_number_filter("Max wet weight [kg]", "wet_weight_kg", data, step=50.0)
 
    st.markdown("**Fuel & emissions**")
    fuel_col, imo_col = st.columns(2)
    with fuel_col:
        maximum_fuel = max_number_filter(
            "Max fuel consumption [g/bkW-h]",
            "fuel_consumption_g_bkwh",
            data,
            step=1.0,
            help_text="Uses the published catalogue value only; not an efficiency recommendation.",
        )
    with imo_col:
        selected_imo = st.multiselect(
            "IMO / emission tier",
            options=imo_options,
            default=[],
            key="selected_imo",
            placeholder="All IMO / emission values",
            help=(
                "Select a NOx tier to include every catalogue value that lists it — "
                "e.g. IMO Tier II also matches rows labelled 'IMO Tier II / IMO Tier III'. "
                "Leave empty to include all published and blank IMO values."
            ),
        )
 
    return {
        "brands": selected_brands,
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
        "selected_imo": selected_imo,
    }
 
 
def model_count(data: pd.DataFrame) -> int:
    """Count distinct brand + model products (rating rows that share a model count once)."""
    if data.empty:
        return 0
    return int(data.groupby(["brand", "model"], sort=False).ngroups)
 
 
def render_results(exact: pd.DataFrame, unverified: pd.DataFrame, missing_columns: list[str]) -> None:
    """Render metrics, export, matches and unverified rows."""
    st.subheader("3. Filter results")
 
    metric_left, metric_right = st.columns(2)
    metric_left.metric(
        "Matching models",
        f"{model_count(exact):,}",
        help="Distinct products that satisfy every applied filter (voltage/frequency variants of the same model count once).",
    )
    metric_right.metric(
        "Unverified models",
        f"{model_count(unverified):,}",
        help="Distinct products that pass the known criteria but have a blank field used by an applied filter.",
    )
 
    if not exact.empty or not unverified.empty:
        try:
            export_bytes = cached_export(exact, st.session_state["filter_values"], unverified)
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
 
    if missing_columns:
        readable = [DISPLAY_NAMES.get(column, column) for column in missing_columns]
        st.warning(
            "The uploaded database does not contain the following enabled filter column(s): "
            + ", ".join(readable)
            + ". No confirmed match can be verified for those filters."
        )
 
    if exact.empty:
        st.warning(
            "No confirmed generator products were found with the selected filters. "
            "Try widening the power range or clearing one or more filters."
        )
    else:
        render_grouped_results(exact)
 
    if not unverified.empty:
        with st.expander(f"Products with missing data for an applied filter ({model_count(unverified):,})", expanded=False):
            st.caption(
                "These rows satisfy the known criteria, but at least one applied filter field is "
                "blank. They are not confirmed matches and require catalogue verification."
            )
            render_grouped_results(unverified)
 
 
def main() -> None:
    st.title(APP_TITLE)
    st.caption(APP_DESCRIPTION)
    st.markdown(f'<div class="small-muted">{SCOPE_NOTE}</div>', unsafe_allow_html=True)
 
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
 
    st.success(
        f'Loaded {stats["usable_rows"]:,} rating rows across '
        f'{stats["brand_count"]:,} brands and {stats["model_count"]:,} models.'
    )
 
    imo_options, imo_option_to_raw = build_imo_options(data)
 
    st.subheader("2. Product filters")
    st.markdown(
        '<div class="small-muted">Every filter is optional — leave a field empty to ignore it.</div>',
        unsafe_allow_html=True,
    )
 
    with st.container(border=True):
        raw_filters = render_filters(data, imo_options)
 
    # Expand tier selections (e.g. "IMO Tier II") into the exact raw catalogue
    # strings they cover (e.g. "IMO II", "IMO Tier II / IMO Tier III"), so matching
    # stays exact and the stored/displayed data is never modified.
    expanded_imo_values = sorted(
        {
            raw
            for option in raw_filters.pop("selected_imo")
            for raw in imo_option_to_raw.get(option, {option})
        }
    )
    filters = {**raw_filters, "imo_values": expanded_imo_values}
 
    find_col, reset_col = st.columns([3, 1])
    find_clicked = find_col.button(
        "Find generator products", type="primary", use_container_width=True
    )
    if reset_col.button("Reset filters", use_container_width=True):
        reset_results(clear_filter_widgets=True)
        st.rerun()
 
    if find_clicked:
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
        st.caption("Results will appear after you click Find generator products.")
        return
 
    render_results(
        st.session_state["filter_result"],
        st.session_state["filter_unverified"],
        st.session_state.get("missing_filter_columns", []),
    )
 
 
if __name__ == "__main__":
    main()
