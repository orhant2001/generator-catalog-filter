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
    /* Title accent bar */
    .app-header {
        border-left: 6px solid #1D4ED8;
        padding: 0.1rem 0 0.1rem 0.9rem;
        margin-bottom: 0.2rem;
    }
    .app-header h1 { margin: 0; font-size: 1.9rem; color: #0B2545; }
    /* Numbered step headers */
    .section-head {
        display: flex;
        align-items: center;
        gap: 0.65rem;
        margin: 1.7rem 0 0.5rem 0;
    }
    .section-num {
        display: inline-flex;
        align-items: center;
        justify-content: center;
        width: 1.9rem;
        height: 1.9rem;
        border-radius: 50%;
        background: #1D4ED8;
        color: #FFFFFF;
        font-weight: 700;
        font-size: 0.95rem;
        flex: 0 0 auto;
    }
    .section-title { color: #0B2545; font-size: 1.3rem; font-weight: 700; }
    /* Colored result stat cards */
    .stat-card {
        display: flex;
        gap: 0.9rem;
        align-items: center;
        border: 1px solid;
        border-radius: 14px;
        padding: 0.95rem 1.15rem;
    }
    .stat-icon { font-size: 1.7rem; line-height: 1; font-weight: 800; flex: 0 0 auto; }
    .stat-value { font-size: 2rem; font-weight: 800; line-height: 1.05; }
    .stat-label { font-weight: 700; color: #0B2545; font-size: 0.95rem; }
    .stat-sub { color: #6B7280; font-size: 0.82rem; margin-top: 0.1rem; }
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
    return (
        f"{numeric:,.0f}"
        if numeric.is_integer()
        else f"{numeric:,.2f}".rstrip("0").rstrip(".")
    )


def section_header(number: str, title: str) -> None:
    """Render a numbered, badge-style step header."""
    st.markdown(
        f'<div class="section-head">'
        f'<span class="section-num">{number}</span>'
        f'<span class="section-title">{title}</span>'
        f"</div>",
        unsafe_allow_html=True,
    )


_STAT_PALETTE = {
    # variant: (accent color, background, border, icon)
    "match": ("#059669", "#ECFDF5", "#A7F3D0", "✓"),
    "warn": ("#D97706", "#FFFBEB", "#FDE68A", "⚠"),
    "muted": ("#6B7280", "#F3F4F6", "#E5E7EB", "—"),
}


def stat_card(container: Any, value: str, label: str, subtitle: str, variant: str) -> None:
    """Render a colored at-a-glance stat card into the given column/container."""
    color, background, border, icon = _STAT_PALETTE[variant]
    container.markdown(
        f'<div class="stat-card" style="background:{background}; border-color:{border};">'
        f'<div class="stat-icon" style="color:{color};">{icon}</div>'
        f"<div>"
        f'<div class="stat-value" style="color:{color};">{value}</div>'
        f'<div class="stat-label">{label}</div>'
        f'<div class="stat-sub">{subtitle}</div>'
        f"</div></div>",
        unsafe_allow_html=True,
    )


# ---------------------------------------------------------------------------
# Emission-standard classification
# ---------------------------------------------------------------------------
#
# The Excel column is still named "IMO", and the downstream filter still uses
# the existing "imo_values" key. Only the option-building logic changes.
#
# A single raw cell may belong to multiple accepted classes:
#
#     IMO Tier III / EPA 4 (COM) / Stage V
#
# is mapped to:
#
#     IMO Tier III
#     EPA Tier 4
#     EU Stage V
#
# Parenthetical notes such as (COM), (gas mode) and (SCR) are ignored only
# while classifying. The original Excel text remains unchanged in results and
# exports.

_EMISSION_CLASS_ORDER = (
    "IMO_I",
    "IMO_II",
    "IMO_III",
    "EPA_2",
    "EPA_3",
    "EPA_4",
    "EU_STAGE_II",
    "EU_STAGE_IIIA",
    "EU_STAGE_IIIB",
    "EU_STAGE_IV",
    "EU_STAGE_V",
    "NST",
)

_EMISSION_LABELS = {
    "IMO_I": "IMO Tier I",
    "IMO_II": "IMO Tier II",
    "IMO_III": "IMO Tier III",
    "EPA_2": "EPA Tier 2",
    "EPA_3": "EPA Tier 3",
    "EPA_4": "EPA Tier 4",
    "EU_STAGE_II": "EU Stage II",
    "EU_STAGE_IIIA": "EU Stage IIIA",
    "EU_STAGE_IIIB": "EU Stage IIIB",
    "EU_STAGE_IV": "EU Stage IV",
    "EU_STAGE_V": "EU Stage V",
    "NST": "IMO kapsamı dışında (NST)",
}

_PARENTHETICAL_NOTE_RE = re.compile(r"\([^)]*\)")

_IMO_RE = re.compile(
    r"\bIMO\b"
    r"\s*(?:NOX\s*)?"
    r"(?:TIER\s*)?"
    r"(?P<first>III|II|I|3|2|1)\b"
    r"(?:"
    r"\s*(?:/|&|\+|,|\bAND\b|\bVE\b)\s*"
    r"(?:IMO\s*)?"
    r"(?:NOX\s*)?"
    r"(?:TIER\s*)?"
    r"(?P<second>III|II|I|3|2|1)\b"
    r")?",
    re.IGNORECASE,
)

_IMO_TOKEN_TO_CLASS = {
    "I": "IMO_I",
    "1": "IMO_I",
    "II": "IMO_II",
    "2": "IMO_II",
    "III": "IMO_III",
    "3": "IMO_III",
}

_EPA_RE = re.compile(
    r"\b(?:U\.?S\.?\s*)?EPA\b"
    r"\s*(?:TIER\s*)?"
    r"(?P<tier>IV|III|II|4|3|2)\b",
    re.IGNORECASE,
)

_EPA_TOKEN_TO_CLASS = {
    "II": "EPA_2",
    "2": "EPA_2",
    "III": "EPA_3",
    "3": "EPA_3",
    "IV": "EPA_4",
    "4": "EPA_4",
}

_STAGE_RE = re.compile(
    r"\b(?:EU\s*)?STAGE\s*"
    r"(?P<stage>IIIB|IIIA|IV|III|II|V|5|4|3B|3A|3|2)\b",
    re.IGNORECASE,
)

_STAGE_TOKEN_TO_CLASS = {
    "II": "EU_STAGE_II",
    "2": "EU_STAGE_II",
    "IIIA": "EU_STAGE_IIIA",
    "3A": "EU_STAGE_IIIA",
    "III": "EU_STAGE_IIIA",
    "3": "EU_STAGE_IIIA",
    "IIIB": "EU_STAGE_IIIB",
    "3B": "EU_STAGE_IIIB",
    "IV": "EU_STAGE_IV",
    "4": "EU_STAGE_IV",
    "V": "EU_STAGE_V",
    "5": "EU_STAGE_V",
}

_NST_RE = re.compile(
    r"\bNST\b"
    r"|\bNOT\s+SUBJECT\s+TO\b"
    r"|\bIMO\s+KAPSAMI(?:NA)?\s+DIŞINDA\b"
    r"|\bIMO\s+KAPSAMI(?:NA)?\s+(?:TABİ|DAHİL)\s+DEĞİL\b",
    re.IGNORECASE,
)


def emission_classes_in(value: Any) -> list[str]:
    """Return every accepted emission class contained in a raw catalogue value.

    Examples
    --------
    "IMO Tier III / EPA 4 (COM) / Stage V"
        -> ["IMO_III", "EPA_4", "EU_STAGE_V"]
    "IMO Tier II / IMO Tier III (gas mode)"
        -> ["IMO_II", "IMO_III"]

    Parenthetical text is ignored for classification. Unrecognized text does
    not become a separate filter option.
    """
    if value is None or pd.isna(value):
        return []
    text = str(value).strip().upper()
    if not text:
        return []
    # Ignore notes such as (COM), (gas mode) and (SCR) for classification.
    text = _PARENTHETICAL_NOTE_RE.sub(" ", text)

    detected: set[str] = set()

    for match in _IMO_RE.finditer(text):
        for group_name in ("first", "second"):
            token = match.group(group_name)
            if token:
                emission_class = _IMO_TOKEN_TO_CLASS.get(token.upper())
                if emission_class:
                    detected.add(emission_class)

    for match in _EPA_RE.finditer(text):
        emission_class = _EPA_TOKEN_TO_CLASS.get(match.group("tier").upper())
        if emission_class:
            detected.add(emission_class)

    for match in _STAGE_RE.finditer(text):
        emission_class = _STAGE_TOKEN_TO_CLASS.get(match.group("stage").upper())
        if emission_class:
            detected.add(emission_class)

    if _NST_RE.search(text):
        detected.add("NST")

    return [
        emission_class
        for emission_class in _EMISSION_CLASS_ORDER
        if emission_class in detected
    ]


def build_imo_options(data: pd.DataFrame) -> tuple[list[str], dict[str, set[str]]]:
    """Build canonical emission options mapped to their exact raw Excel values.

    The function name is kept for compatibility with the rest of the existing
    application. A raw value may be mapped to more than one option.

    Example
    -------
    Raw value:
        "IMO Tier III / EPA 4 (COM) / Stage V"
    Mappings:
        "IMO Tier III" -> raw value
        "EPA Tier 4"   -> raw value
        "EU Stage V"   -> raw value
    """
    raw_values = text_options(data, "imo")

    class_to_raw: dict[str, set[str]] = {
        emission_class: set()
        for emission_class in _EMISSION_CLASS_ORDER
    }
    for raw_value in raw_values:
        for emission_class in emission_classes_in(raw_value):
            class_to_raw[emission_class].add(raw_value)

    options: list[str] = []
    option_to_raw: dict[str, set[str]] = {}
    for emission_class in _EMISSION_CLASS_ORDER:
        matching_raw_values = class_to_raw[emission_class]
        if not matching_raw_values:
            continue
        label = _EMISSION_LABELS[emission_class]
        options.append(label)
        option_to_raw[label] = matching_raw_values

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
                lambda value: (
                    "Not available"
                    if pd.isna(value)
                    else f"{float(value):,.{places}f}"
                )
            )
    for column in table.columns:
        if column not in decimal_places:
            table[column] = (
                table[column]
                .astype(object)
                .where(table[column].notna(), "Not available")
            )
    return table


def render_grouped_results(data: pd.DataFrame, accent: str = "match") -> None:
    """Show one collapsible box per model; each box lists that model's rating rows.

    Rows that share a model but differ only by voltage/frequency (e.g. PP16V4000P63
    listed several times) are collapsed under a single expandable entry. When the box
    is opened, each voltage/frequency variation is shown on its own line.
    """
    grouped = data.groupby(["brand", "model"], sort=False)
    caption = f"{grouped.ngroups:,} model · {len(data):,} rating row(s)"
    if accent == "warn":
        st.markdown(
            f'<div style="color:#D97706; font-size:0.9rem; margin-bottom:0.2rem;">'
            f"{caption}</div>",
            unsafe_allow_html=True,
        )
    else:
        st.caption(caption)

    marker = "⚠ " if accent == "warn" else ""
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
        header = f"{marker}{brand} — {model} · {count} {label}{power_text}"
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
    "power_min_input",
    "power_max_input",
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
    for key in (
        "filter_result",
        "filter_unverified",
        "filter_values",
        "missing_filter_columns",
    ):
        st.session_state.pop(key, None)
    if clear_filter_widgets:
        for key in FILTER_WIDGET_KEYS:
            st.session_state.pop(key, None)


def reset_filters_callback() -> None:
    """Reset every filter widget from a button ``on_click`` callback.

    Running the clear inside a callback (rather than inline after the widgets
    are drawn) means it executes *before* any widget is instantiated on the
    next run, so the widgets always rebuild from their defaults. This avoids
    the browser occasionally restoring a stale value when a widget declares
    both a ``default``/``value`` and a ``key``.
    """
    reset_results(clear_filter_widgets=True)


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
        power_step = max(round((power_max - power_min) / 100, 1), 0.1)

        # Slider ve sayı kutuları aynı değeri paylaşsın diye ortak state.
        if "power_range" not in st.session_state:
            st.session_state["power_range"] = (power_min, power_max)
        if "power_min_input" not in st.session_state:
            st.session_state["power_min_input"] = float(power_min)
        if "power_max_input" not in st.session_state:
            st.session_state["power_max_input"] = float(power_max)

        def _clamp_power(value: float) -> float:
            return min(max(float(value), power_min), power_max)

        def _power_from_slider() -> None:
            low, high = st.session_state["power_range"]
            st.session_state["power_min_input"] = float(low)
            st.session_state["power_max_input"] = float(high)

        def _power_from_inputs() -> None:
            low = _clamp_power(st.session_state["power_min_input"])
            high = _clamp_power(st.session_state["power_max_input"])
            if low > high:
                low, high = high, low
            st.session_state["power_range"] = (low, high)
            st.session_state["power_min_input"] = low
            st.session_state["power_max_input"] = high

        st.slider(
            "Power range [kW]",
            min_value=power_min,
            max_value=power_max,
            step=power_step,
            key="power_range",
            on_change=_power_from_slider,
            help=(
                "Drag the handles or type exact values below. "
                "Full range means no power filter."
            ),
        )

        p_min_col, p_max_col = st.columns(2)
        with p_min_col:
            st.number_input(
                "Min power [kW]",
                min_value=power_min,
                max_value=power_max,
                step=power_step,
                key="power_min_input",
                on_change=_power_from_inputs,
                help="Type an exact minimum — stays in sync with the slider.",
            )
        with p_max_col:
            st.number_input(
                "Max power [kW]",
                min_value=power_min,
                max_value=power_max,
                step=power_step,
                key="power_max_input",
                on_change=_power_from_inputs,
                help="Type an exact maximum — stays in sync with the slider.",
            )

        power_low, power_high = st.session_state["power_range"]
        minimum_power = power_low if power_low > power_min else None
        maximum_power = power_high if power_high < power_max else None

    st.markdown("##### Size, weight, fuel & emissions")
    st.caption(
        "Enter a maximum to apply a limit, or leave a field blank to ignore it. "
        "Products with a blank value for an applied limit are listed separately "
        "as unverified."
    )

    st.markdown("**Maximum dimensions [mm]**")
    dim1, dim2, dim3 = st.columns(3)
    with dim1:
        maximum_width = max_number_filter(
            "Max width [mm]", "width_mm", data, step=50.0
        )
    with dim2:
        maximum_depth = max_number_filter(
            "Max depth [mm]", "depth_mm", data, step=50.0
        )
    with dim3:
        maximum_height = max_number_filter(
            "Max height [mm]", "height_mm", data, step=50.0
        )

    st.markdown("**Maximum weight [kg]**")
    weight1, weight2 = st.columns(2)
    with weight1:
        maximum_dry_weight = max_number_filter(
            "Max dry weight [kg]", "dry_weight_kg", data, step=50.0
        )
    with weight2:
        maximum_wet_weight = max_number_filter(
            "Max wet weight [kg]", "wet_weight_kg", data, step=50.0
        )

    st.markdown("**Fuel & emissions**")
    fuel_col, imo_col = st.columns(2)
    with fuel_col:
        maximum_fuel = max_number_filter(
            "Max fuel consumption [g/bkW-h]",
            "fuel_consumption_g_bkwh",
            data,
            step=1.0,
            help_text=(
                "Uses the published catalogue value only; "
                "not an efficiency recommendation."
            ),
        )
    with imo_col:
        selected_imo = st.multiselect(
            "Emission standard",
            options=imo_options,
            default=[],
            key="selected_imo",
            placeholder="All recognized emission standards",
            help=(
                "A catalogue value may match more than one standard. "
                "For example, 'IMO Tier III / EPA 4 (COM) / Stage V' "
                "matches IMO Tier III, EPA Tier 4 and EU Stage V. "
                "Parenthetical notes such as COM, gas mode and SCR are not "
                "separate filter classes. Selecting multiple options uses OR logic."
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
    """Count distinct brand + model products (rating rows sharing a model count once)."""
    if data.empty:
        return 0
    return int(data.groupby(["brand", "model"], sort=False).ngroups)


SORT_MODES = (
    "Match order",
    "Power (low → high)",
    "Power (high → low)",
    "Brand / model (A–Z)",
)


def sort_for_display(data: pd.DataFrame, mode: str) -> pd.DataFrame:
    """Reorder rating rows for display; groups then follow first-appearance order."""
    if data.empty or mode == "Match order":
        return data
    if mode == "Power (low → high)":
        return data.sort_values("power_kw", kind="stable")
    if mode == "Power (high → low)":
        return data.sort_values("power_kw", ascending=False, kind="stable")
    if mode == "Brand / model (A–Z)":
        return data.sort_values(["brand", "model"], kind="stable")
    return data


def render_results(
    exact: pd.DataFrame,
    unverified: pd.DataFrame,
    missing_columns: list[str],
) -> None:
    """Render metrics, export, matches and unverified rows."""
    section_header("3", "Filter results")

    match_models = model_count(exact)
    unverified_models = model_count(unverified)

    metric_left, metric_right = st.columns(2)
    stat_card(
        metric_left,
        f"{match_models:,}",
        "Matching models",
        "Satisfy every applied filter" if match_models else "No products matched",
        "match" if match_models else "muted",
    )
    stat_card(
        metric_right,
        f"{unverified_models:,}",
        "Unverified models",
        "Need catalogue verification"
        if unverified_models
        else "None — all matches verified",
        "warn" if unverified_models else "muted",
    )

    st.write("")

    if not exact.empty or not unverified.empty:
        try:
            export_bytes = cached_export(
                exact,
                st.session_state["filter_values"],
                unverified,
            )
        except Exception:
            st.error("The filtered Excel export could not be generated.")
        else:
            st.download_button(
                "⬇ Export filtered results to Excel",
                data=export_bytes,
                file_name="filtered_generator_catalog_results.xlsx",
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),
                use_container_width=True,
            )

    if missing_columns:
        readable = [DISPLAY_NAMES.get(column, column) for column in missing_columns]
        st.warning(
            "The uploaded database does not contain the following enabled "
            "filter column(s): "
            + ", ".join(readable)
            + ". No confirmed match can be verified for those filters."
        )

    if not exact.empty or not unverified.empty:
        sort_col, _ = st.columns([2, 3])
        sort_mode = sort_col.selectbox(
            "Sort models by",
            SORT_MODES,
            key="results_sort",
        )
        exact = sort_for_display(exact, sort_mode)
        unverified = sort_for_display(unverified, sort_mode)

    if exact.empty:
        st.warning(
            "No confirmed generator products were found with the selected filters. "
            "Try widening the power range or clearing one or more filters."
        )
    else:
        st.markdown(
            '<div style="color:#059669; font-weight:600; '
            'margin:0.6rem 0 0.2rem 0;">'
            "✓ Confirmed matches — open a model to see its voltage / "
            "frequency ratings.</div>",
            unsafe_allow_html=True,
        )
        render_grouped_results(exact, accent="match")

    if not unverified.empty:
        st.markdown(
            '<div style="color:#D97706; font-weight:600; '
            'margin:1.2rem 0 0.2rem 0;">'
            f"⚠ {unverified_models:,} unverified model(s) — pass the known "
            "criteria, but a field used by an applied filter is blank.</div>",
            unsafe_allow_html=True,
        )
        if st.toggle(
            "Show unverified models",
            value=False,
            key="show_unverified",
        ):
            render_grouped_results(unverified, accent="warn")


def main() -> None:
    st.markdown(
        f'<div class="app-header"><h1>⚙️ {APP_TITLE}</h1></div>',
        unsafe_allow_html=True,
    )
    st.caption(APP_DESCRIPTION)
    st.markdown(
        f'<div class="small-muted">{SCOPE_NOTE}</div>',
        unsafe_allow_html=True,
    )

    section_header("1", "Upload generator database")
    uploaded_file = st.file_uploader(
        'Upload an .xlsx workbook containing the worksheet "Jeneratör Verileri".',
        type=["xlsx"],
        accept_multiple_files=False,
    )

    if uploaded_file is None:
        reset_results()
        st.info(
            "Upload the prepared generator Excel database to activate "
            "the catalogue filters."
        )
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
        st.error(
            "The Excel database could not be processed due to an unexpected error."
        )
        return

    st.success(
        f'Loaded {stats["usable_rows"]:,} rating rows across '
        f'{stats["brand_count"]:,} brands and {stats["model_count"]:,} models.'
    )

    imo_options, imo_option_to_raw = build_imo_options(data)

    section_header("2", "Product filters")
    st.markdown(
        '<div class="small-muted">'
        "Every filter is optional — leave a field empty to ignore it."
        "</div>",
        unsafe_allow_html=True,
    )
    with st.container(border=True):
        raw_filters = render_filters(data, imo_options)

    # Expand each canonical emission selection into the exact raw catalogue
    # strings that contain it. One raw string may belong to multiple classes.
    #
    # Example:
    #     "IMO Tier III / EPA 4 (COM) / Stage V"
    # is included when the user selects IMO Tier III, EPA Tier 4 or EU Stage V.
    #
    # The downstream data filter remains an exact text match and the stored
    # or displayed Excel data is never modified.
    expanded_imo_values = sorted(
        {
            raw_value
            for selected_option in raw_filters.pop("selected_imo")
            for raw_value in imo_option_to_raw.get(
                selected_option,
                {selected_option},
            )
        }
    )
    filters = {**raw_filters, "imo_values": expanded_imo_values}

    find_col, reset_col = st.columns([3, 1])
    find_clicked = find_col.button(
        "Find generator products",
        type="primary",
        use_container_width=True,
    )
    reset_col.button(
        "Reset filters",
        use_container_width=True,
        on_click=reset_filters_callback,
    )

    if find_clicked:
        try:
            exact, unverified, missing_columns = filter_catalog(data, filters)
        except ValueError as exc:
            st.error(str(exc))
        except Exception:
            st.error(
                "The selected filters could not be applied due to an unexpected error."
            )
        else:
            st.session_state["filter_result"] = exact
            st.session_state["filter_unverified"] = unverified
            st.session_state["filter_values"] = filters
            st.session_state["missing_filter_columns"] = missing_columns

    if "filter_result" not in st.session_state:
        st.caption(
            "Results will appear after you click Find generator products."
        )
        return

    if st.session_state.get("filter_values") != filters:
        st.warning(
            "Filters have changed since these results — click "
            "**Find generator products** to refresh."
        )

    render_results(
        st.session_state["filter_result"],
        st.session_state["filter_unverified"],
        st.session_state.get("missing_filter_columns", []),
    )


if __name__ == "__main__":
    main()
