"""
Productivity Report Analyzer – Streamlit UI.

Upload one or more Productivity Reports and a shared Standard effort table,
calculate task-level productivity per file and combined, browse results with
a report slider, and export each report as its own Excel workbook.
"""

from __future__ import annotations

import io
import zipfile
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st

from src.batch_processor import (
    BatchResult,
    FileInput,
    calculate_batch,
    per_file_summary_table,
    validate_mapped_columns,
)
from src.column_mapper import auto_map_columns, validate_mapping
from src.configuration import (
    PRODUCTIVITY_REQUIRED_COLUMNS,
    STANDARD_REQUIRED_COLUMNS,
    AppConfig,
    validate_coefficients,
)
from src.excel_reader import ExcelReadError, inspect_workbook, list_columns, read_table
from src.productivity_calculator import CalculationResult, round_for_display
from src.report_generator import generate_report_bytes
from src.validators import build_standard_lookup

st.set_page_config(
    page_title="Productivity Report Analyzer",
    page_icon="📊",
    layout="wide",
)

# ---------------------------------------------------------------------------
# Session helpers
# ---------------------------------------------------------------------------

SESSION_DEFAULTS = {
    "prod_files": [],  # list[{"name", "bytes"}]
    "std_bytes": None,
    "std_name": None,
    "tpl_bytes": None,
    "tpl_name": None,
    "batch_result": None,
    "standard_lookup": None,
    "config_coeffs": None,
    "report_slider_index": 0,
}


def _init_session() -> None:
    for key, value in SESSION_DEFAULTS.items():
        if key not in st.session_state:
            st.session_state[key] = value if key != "prod_files" else []


# ---------------------------------------------------------------------------
# Shared UI helpers
# ---------------------------------------------------------------------------

def _status_badge(value: int) -> None:
    """Render a colored status pill: green OK, yellow Needs review."""
    if value == 0:
        label, bg, fg, border = "OK", "#E7F6ED", "#1B5E20", "#A5D6A7"
    else:
        label, bg, fg, border = "Needs review", "#FFF8E1", "#E65100", "#FFE082"
    st.html(
        f"""
        <div style="
            display:inline-block;
            margin-top:0.15rem;
            padding:0.2rem 0.65rem;
            border-radius:999px;
            font-size:0.78rem;
            font-weight:700;
            letter-spacing:0.02em;
            background:{bg};
            color:{fg};
            border:1px solid {border};
        ">{label}</div>
        """
    )


def _metric_with_status(label: str, value: int) -> None:
    """Metric plus highlighted OK / Needs review status."""
    st.metric(label, value)
    _status_badge(value)


def render_summary_panels(
    result: CalculationResult,
    lookup,
    config: AppConfig,
    *,
    title_prefix: str = "",
    show_effort_totals: bool = True,
) -> None:
    """Render Match overview + Calculation summary using native Streamlit metrics."""
    s = result.summary
    dup_keys = len(lookup.duplicate_identical_keys) if lookup is not None else 0
    amb_keys = len(lookup.ambiguous_keys) if lookup is not None else 0
    unmatched = s.unmatched_task_keys
    invalid_effort = s.invalid_actual_effort + s.zero_actual_effort
    match_title = f"{title_prefix}Match overview".strip()
    calc_title = f"{title_prefix}Calculation summary".strip()

    st.subheader(match_title)
    m1, m2, m3 = st.columns(3)
    with m1:
        _metric_with_status("Duplicate identical standard keys", dup_keys)
    with m2:
        _metric_with_status("Ambiguous standard keys", amb_keys)
    with m3:
        _metric_with_status("Unmatched task keys", unmatched)

    st.subheader(calc_title)
    st.caption("Coverage")
    c1, c2 = st.columns(2)
    c1.metric("Total source tasks", s.total_tasks)
    c2.metric("Valid productivity tasks", s.valid_productivity_tasks)

    st.caption("Validation flags")
    v1, v2, v3 = st.columns(3)
    with v1:
        _metric_with_status("Missing standard", s.missing_standard)
    with v2:
        _metric_with_status("Ambiguous standard", s.ambiguous_standard)
    with v3:
        _metric_with_status("Invalid actual effort", invalid_effort)

    if show_effort_totals:
        actual_md = round(s.total_actual_effort, config.effort_decimals)
        standard_md = round(s.total_standard_effort, config.effort_decimals)
        st.caption("Effort totals")
        e1, e2 = st.columns(2)
        with e1:
            st.info(f"**Total actual effort (valid)**  \n### {actual_md} MD")
        with e2:
            st.success(f"**Total standard effort (valid)**  \n### {standard_md} MD")


def _clear_browse_filter_state() -> None:
    """Drop browse-filter widget keys so the next report starts unfiltered."""
    for key in list(st.session_state.keys()):
        if isinstance(key, str) and key.startswith("browse_"):
            del st.session_state[key]


def _apply_filters(detail: pd.DataFrame, key_prefix: str = "flt") -> Tuple[pd.DataFrame, Dict[str, Any]]:
    """Render filter widgets and return filtered detail."""
    st.subheader("Filters")
    f1, f2, f3 = st.columns(3)
    employees = sorted({e for e in detail["Employee Name"].dropna().unique()})
    complexities = sorted({c for c in detail["Complexity"].dropna().unique()})
    task_types = sorted({t for t in detail["Task Type"].dropna().unique()})
    changes = sorted({c for c in detail["Change/New"].dropna().unique()})
    levels = sorted({lv for lv in detail["Professional Level"].dropna().unique()})
    statuses = sorted(detail["Validation Status"].dropna().unique())
    source_files = (
        sorted({f for f in detail["Source File"].dropna().unique()})
        if "Source File" in detail.columns
        else []
    )

    # Drop stale selections that are not in the current report's options
    # (common after Next/Previous when widget keys are reused).
    def _sanitize_multiselect(state_key: str, options: list) -> None:
        if state_key in st.session_state:
            current = st.session_state[state_key]
            if isinstance(current, list):
                st.session_state[state_key] = [v for v in current if v in options]

    _sanitize_multiselect(f"{key_prefix}_emp", employees)
    _sanitize_multiselect(f"{key_prefix}_cx", complexities)
    _sanitize_multiselect(f"{key_prefix}_tt", task_types)
    _sanitize_multiselect(f"{key_prefix}_ch", changes)
    _sanitize_multiselect(f"{key_prefix}_lv", levels)
    _sanitize_multiselect(f"{key_prefix}_st", statuses)
    _sanitize_multiselect(f"{key_prefix}_src", source_files)

    sel_emp = f1.multiselect("Employee Name", employees, key=f"{key_prefix}_emp")
    sel_cx = f2.multiselect("Complexity", complexities, key=f"{key_prefix}_cx")
    sel_tt = f3.multiselect("Task Type", task_types, key=f"{key_prefix}_tt")

    f4, f5, f6 = st.columns(3)
    sel_ch = f4.multiselect("Change/New", changes, key=f"{key_prefix}_ch")
    sel_lv = f5.multiselect("Professional Level", levels, key=f"{key_prefix}_lv")
    sel_st = f6.multiselect("Validation Status", statuses, key=f"{key_prefix}_st")

    sel_src: List[str] = []
    # Source File filter is only useful when combining multiple reports.
    if len(source_files) > 1:
        sel_src = st.multiselect("Source File", source_files, key=f"{key_prefix}_src")

    date_min = detail["Finished Date"].min()
    date_max = detail["Finished Date"].max()
    date_range = None
    if pd.notna(date_min) and pd.notna(date_max):
        min_date = pd.Timestamp(date_min).date()
        max_date = pd.Timestamp(date_max).date()
        date_key = f"{key_prefix}_dates"
        # Reset stale date range that does not overlap this report's dates.
        if date_key in st.session_state:
            raw_state = st.session_state[date_key]
            if isinstance(raw_state, (list, tuple)) and len(raw_state) == 2:
                try:
                    s0 = pd.Timestamp(raw_state[0]).date()
                    s1 = pd.Timestamp(raw_state[1]).date()
                    if s1 < min_date or s0 > max_date:
                        del st.session_state[date_key]
                except Exception:  # noqa: BLE001
                    del st.session_state[date_key]
            else:
                del st.session_state[date_key]

        raw = st.date_input(
            "Finished Date range",
            value=(min_date, max_date),
            key=date_key,
        )
        if isinstance(raw, (list, tuple)) and len(raw) == 2:
            d1, d2 = raw
            date_range = (
                pd.Timestamp(d1).normalize(),
                pd.Timestamp(d2).normalize() + pd.Timedelta(days=1),
            )

    filtered = detail
    before_count = len(filtered)
    if sel_emp:
        filtered = filtered[filtered["Employee Name"].isin(sel_emp)]
    if sel_cx:
        filtered = filtered[filtered["Complexity"].isin(sel_cx)]
    if sel_tt:
        filtered = filtered[filtered["Task Type"].isin(sel_tt)]
    if sel_ch:
        filtered = filtered[filtered["Change/New"].isin(sel_ch)]
    if sel_lv:
        filtered = filtered[filtered["Professional Level"].isin(sel_lv)]
    if sel_st:
        filtered = filtered[filtered["Validation Status"].isin(sel_st)]
    if sel_src and "Source File" in filtered.columns:
        filtered = filtered[filtered["Source File"].isin(sel_src)]
    if date_range is not None and not filtered.empty:
        dates = pd.to_datetime(filtered["Finished Date"], errors="coerce")
        if getattr(dates.dt, "tz", None) is not None:
            dates = dates.dt.tz_localize(None)
        mask = (dates >= date_range[0]) & (dates < date_range[1])
        # If a stale date widget would hide every row, keep the rows instead.
        if mask.any() or before_count == 0:
            filtered = filtered[mask]
    return filtered, {}


def _build_comparison_frame(filtered: pd.DataFrame) -> pd.DataFrame:
    compare_cols = [
        c
        for c in [
            "Source File",
            "Source Row",
            "Task Key",
            "Employee Name",
            "Finished Date",
            "Task Type",
            "Change/New",
            "Complexity",
            "Total Effort (MD)",
            "Standard Effort",
        ]
        if c in filtered.columns
    ]
    compare = filtered[compare_cols].copy()
    if "Total Effort (MD)" in compare.columns:
        compare = compare.rename(columns={"Total Effort (MD)": "EffortManday"})
    return compare


def _highlight_effort_diff(row: pd.Series) -> list:
    styles = [""] * len(row)
    if "EffortManday" not in row.index or "Standard Effort" not in row.index:
        return styles
    actual = row["EffortManday"]
    standard = row["Standard Effort"]
    if (
        actual is None
        or standard is None
        or (isinstance(actual, float) and pd.isna(actual))
        or (isinstance(standard, float) and pd.isna(standard))
    ):
        return styles
    idx = list(row.index).index("EffortManday")
    if float(actual) == float(standard):
        styles[idx] = "background-color: #C8E6C9; font-weight: 600"
    else:
        styles[idx] = "background-color: #FFE566; font-weight: 600"
    return styles


def _build_calculated_frame(filtered: pd.DataFrame, config: AppConfig) -> pd.DataFrame:
    has_validation_issues = (
        "Validation Status" in filtered.columns
        and (filtered["Validation Status"] != "OK").any()
    )
    show_cols = [
        c
        for c in [
            "Source File",
            "Source Row",
            "Task Key",
            "Employee Name",
            "Professional Level",
            "Finished Date",
            "Task Type",
            "Change/New",
            "Complexity",
            "Total Effort (MD)",
            "Standard Effort",
            "Productivity",
            "Effort For Specialist",
            *(
                ["Validation Status", "Validation Issues"]
                if has_validation_issues
                else []
            ),
        ]
        if c in filtered.columns
    ]
    display = filtered[show_cols].copy()
    if "Productivity" in display.columns:
        decimals = config.productivity_decimals

        def _fmt_productivity(value):
            if value is None or (isinstance(value, float) and pd.isna(value)):
                return None
            if isinstance(value, (int, float)):
                return f"{value:.{decimals}f}%"
            return value

        display["Productivity"] = display["Productivity"].apply(_fmt_productivity)
    return display


def _render_comparison_and_detail(filtered: pd.DataFrame, config: AppConfig) -> None:
    """Single grouped section: Productivity Report + Calculated results."""
    with st.container(border=True):
        st.subheader("Task details")
        st.caption(
            "Compare source effort against the standard, then review calculated productivity. "
            "EffortManday: green = matches Standard Effort, yellow = differs."
        )

        tab_compare, tab_calc = st.tabs(
            ["Productivity Report", "Calculated results"]
        )

        with tab_compare:
            compare = _build_comparison_frame(filtered)
            if compare.empty:
                st.info("No rows match the current filters.")
            else:
                effort_fmt = f"{{:.{config.effort_decimals}f}}"
                formatters = {
                    col: effort_fmt
                    for col in ("EffortManday", "Standard Effort")
                    if col in compare.columns
                }
                styled = compare.style.apply(_highlight_effort_diff, axis=1).format(
                    formatters, na_rep=""
                )
                st.dataframe(styled, use_container_width=True, hide_index=True)

        with tab_calc:
            display = _build_calculated_frame(filtered, config)
            if display.empty:
                st.info("No rows match the current filters.")
            else:
                st.dataframe(display, use_container_width=True, hide_index=True)


# ---------------------------------------------------------------------------
# UI sections
# ---------------------------------------------------------------------------

def section_upload() -> None:
    st.header("A. Upload files")
    st.caption(
        "Upload one or more Productivity Reports and a shared Standard effort table. "
        "An Output Template is optional. CSV does not preserve Excel formatting or worksheets. "
        "Macro-enabled (.xlsm) files are read as data only — macros are never executed."
    )

    col1, col2, col3 = st.columns(3)
    with col1:
        prod_files = st.file_uploader(
            "Productivity Report(s)",
            type=["xlsx", "xlsm", "csv"],
            accept_multiple_files=True,
            key="prod_uploader",
            help="Upload one or many employee task reports.",
        )
    with col2:
        std_file = st.file_uploader(
            "Standard (fixed reference)",
            type=["xlsx", "xlsm", "csv"],
            key="std_uploader",
            help="Shared standard effort table for all reports. Upload once; kept in session.",
        )
    with col3:
        tpl_file = st.file_uploader(
            "Output Template (optional)",
            type=["xlsx", "xlsm"],
            key="tpl_uploader",
        )

    # Only replace session uploads / clear results when file content actually changes.
    # File uploaders remain non-None on every rerun; clearing batch_result each time
    # would wipe calculations when clicking Next / Previous.
    if prod_files:
        new_files = [{"name": f.name, "bytes": f.getvalue()} for f in prod_files]
        prev = st.session_state.prod_files or []
        prev_sig = [(p["name"], len(p["bytes"])) for p in prev]
        new_sig = [(p["name"], len(p["bytes"])) for p in new_files]
        if prev_sig != new_sig:
            st.session_state.prod_files = new_files
            st.session_state.batch_result = None
            st.session_state.report_slider_index = 0
            if "report_picker_top" in st.session_state:
                del st.session_state["report_picker_top"]

    if std_file is not None:
        data = std_file.getvalue()
        name = std_file.name
        std_changed = (
            st.session_state.std_name != name
            or st.session_state.std_bytes is None
            or len(st.session_state.std_bytes) != len(data)
        )
        st.session_state.std_bytes = data
        st.session_state.std_name = name
        if std_changed:
            st.session_state.batch_result = None

    if tpl_file is not None:
        data = tpl_file.getvalue()
        name = tpl_file.name
        st.session_state.tpl_bytes = data
        st.session_state.tpl_name = name

    if st.session_state.std_bytes and st.session_state.std_name:
        st.success(
            f"Standard loaded: **{st.session_state.std_name}** (shared across all reports)."
        )
    if st.session_state.prod_files:
        names = ", ".join(f"**{f['name']}**" for f in st.session_state.prod_files)
        st.info(f"Productivity Reports ({len(st.session_state.prod_files)}): {names}")
    if st.session_state.tpl_bytes and st.session_state.tpl_name:
        st.info(f"Output Template: **{st.session_state.tpl_name}**")


def _sheet_selector(label: str, data: bytes, filename: str, key: str) -> str:
    info = inspect_workbook(data, filename=filename)
    if info.has_vba:
        st.warning(
            f"{filename}: macro-enabled workbook detected. Macros will not be executed."
        )
    if info.is_csv:
        return "CSV"
    sheet = st.selectbox(
        label,
        options=info.sheet_names,
        index=info.sheet_names.index(info.suggested_sheet)
        if info.suggested_sheet in info.sheet_names
        else 0,
        key=key,
    )
    return sheet


def section_mapping_and_preview() -> Optional[Dict]:
    st.header("B. Column mapping")
    if not st.session_state.prod_files or not st.session_state.std_bytes:
        st.warning("Upload at least one Productivity Report and a Standard file to continue.")
        return None

    try:
        # Per-file worksheet selection
        prod_frames: List[Tuple[str, str, pd.DataFrame]] = []
        st.subheader("Worksheet selection")
        for i, pf in enumerate(st.session_state.prod_files):
            sheet = _sheet_selector(
                f"Worksheet – {pf['name']}",
                pf["bytes"],
                pf["name"],
                f"prod_sheet_{i}_{pf['name']}",
            )
            df = read_table(
                pf["bytes"],
                sheet_name=None if sheet == "CSV" else sheet,
                filename=pf["name"],
            )
            prod_frames.append((pf["name"], sheet, df))

        std_sheet = _sheet_selector(
            "Standard worksheet",
            st.session_state.std_bytes,
            st.session_state.std_name,
            "std_sheet",
        )
        tpl_sheet = None
        if st.session_state.tpl_bytes:
            tpl_sheet = _sheet_selector(
                "Output Template worksheet",
                st.session_state.tpl_bytes,
                st.session_state.tpl_name,
                "tpl_sheet",
            )

        std_df = read_table(
            st.session_state.std_bytes,
            sheet_name=None if std_sheet == "CSV" else std_sheet,
            filename=st.session_state.std_name,
        )
    except ExcelReadError as exc:
        st.error(str(exc))
        return None

    # Column mapping based on first file; union of headers offered as choices
    first_cols = list_columns(prod_frames[0][2])
    all_prod_cols: List[str] = []
    seen = set()
    for _, _, df in prod_frames:
        for c in list_columns(df):
            if c not in seen:
                seen.add(c)
                all_prod_cols.append(c)
    std_cols = list_columns(std_df)

    auto_prod = auto_map_columns(first_cols, PRODUCTIVITY_REQUIRED_COLUMNS)
    auto_std = auto_map_columns(std_cols, STANDARD_REQUIRED_COLUMNS)

    st.subheader("Productivity Report columns")
    st.caption("Mappings apply to every uploaded report. Each file must contain the mapped headers.")
    prod_map: Dict[str, Optional[str]] = {}
    cols = st.columns(3)
    for i, (role, expected) in enumerate(PRODUCTIVITY_REQUIRED_COLUMNS.items()):
        with cols[i % 3]:
            options = ["— not mapped —"] + all_prod_cols
            default = auto_prod.get(role)
            idx = options.index(default) if default in options else 0
            chosen = st.selectbox(
                f"{expected}",
                options=options,
                index=idx,
                key=f"prod_map_{role}",
                help=f"Logical role: {role}",
            )
            prod_map[role] = None if chosen.startswith("—") else chosen

    st.subheader("Standard columns")
    std_map: Dict[str, Optional[str]] = {}
    cols2 = st.columns(2)
    for i, (role, expected) in enumerate(STANDARD_REQUIRED_COLUMNS.items()):
        with cols2[i % 2]:
            options = ["— not mapped —"] + std_cols
            default = auto_std.get(role)
            idx = options.index(default) if default in options else 0
            chosen = st.selectbox(
                f"{expected}",
                options=options,
                index=idx,
                key=f"std_map_{role}",
                help=f"Logical role: {role}",
            )
            std_map[role] = None if chosen.startswith("—") else chosen

    mapping_errors = validate_mapping(
        prod_map, PRODUCTIVITY_REQUIRED_COLUMNS, "Productivity Report"
    ) + validate_mapping(std_map, STANDARD_REQUIRED_COLUMNS, "Standard")

    # Per-file column presence check
    if not mapping_errors:
        resolved_prod = {k: v for k, v in prod_map.items() if v}
        for name, _, df in prod_frames:
            mapping_errors.extend(validate_mapped_columns(df, resolved_prod, name))

    if mapping_errors:
        for err in mapping_errors:
            st.error(err)
        mapping_ok = False
    else:
        st.success("All required columns are mapped across uploaded reports.")
        mapping_ok = True

    # Section C – preview
    st.header("C. Data preview")
    total_prod_rows = sum(len(df) for _, _, df in prod_frames)
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Productivity Reports", len(prod_frames))
    c2.metric("Total task rows", total_prod_rows)
    c3.metric("Standard rows", len(std_df))
    c4.metric("Mapped columns", len([v for v in prod_map.values() if v]))

    for name, sheet, df in prod_frames:
        with st.expander(f"Preview – {name} ({sheet}, {len(df)} rows)", expanded=False):
            st.dataframe(df.head(20), use_container_width=True)
    with st.expander("Preview – Standard", expanded=False):
        st.dataframe(std_df.head(20), use_container_width=True)

    return {
        "prod_frames": prod_frames,
        "std_df": std_df,
        "prod_map": prod_map,
        "std_map": std_map,
        "mapping_ok": mapping_ok,
        "tpl_sheet": tpl_sheet,
    }


def section_settings() -> AppConfig:
    st.header("D. Calculation settings")
    st.warning(
        "**Calculation formulas (fixed)**\n\n"
        "- `Productivity (%) = Standard Effort ÷ Actual Effort × 100`\n"
        "- `Effort For Specialist (MD) = Coefficient × Actual Effort ÷ 100`"
    )

    if st.session_state.config_coeffs is None:
        st.session_state.config_coeffs = dict(AppConfig().coefficients)

    coeffs = dict(st.session_state.config_coeffs)
    st.subheader("Professional level coefficients")
    edited = {}
    cols = st.columns(len(coeffs) or 1)
    for i, (level, value) in enumerate(list(coeffs.items())):
        with cols[i % max(len(cols), 1)]:
            edited[level] = st.number_input(
                level,
                min_value=0.0,
                value=float(value),
                step=1.0,
                key=f"coeff_{level}",
            )

    coeff_errors = validate_coefficients(edited)
    if coeff_errors:
        for e in coeff_errors:
            st.error(e)
    else:
        st.session_state.config_coeffs = edited

    c1, c2 = st.columns(2)
    effort_decimals = c1.number_input("Effort decimal places", min_value=0, max_value=6, value=2)
    prod_decimals = c2.number_input("Productivity decimal places", min_value=0, max_value=6, value=2)

    return AppConfig(
        coefficients=dict(st.session_state.config_coeffs),
        effort_decimals=int(effort_decimals),
        productivity_decimals=int(prod_decimals),
    )


def section_calculate(ctx: Dict, config: AppConfig) -> None:
    st.header("E. Calculate")
    disabled = not ctx.get("mapping_ok")
    if st.button("Calculate Productivity", type="primary", disabled=disabled):
        with st.spinner("Calculating all reports…"):
            try:
                prod_map = {k: v for k, v in ctx["prod_map"].items() if v}
                std_map = {k: v for k, v in ctx["std_map"].items() if v}
                lookup = build_standard_lookup(ctx["std_df"], std_map)
                file_inputs = [
                    FileInput(name=name, dataframe=df, sheet_name=sheet)
                    for name, sheet, df in ctx["prod_frames"]
                ]
                batch = calculate_batch(file_inputs, prod_map, lookup, config)
                st.session_state.batch_result = batch
                st.session_state.standard_lookup = lookup
                st.session_state.report_slider_index = 0
                if batch.errors:
                    for err in batch.errors:
                        st.error(err)
            except Exception as exc:  # noqa: BLE001
                st.exception(exc)
                return

    batch: Optional[BatchResult] = st.session_state.batch_result
    lookup = st.session_state.standard_lookup
    if batch is None or batch.combined is None or lookup is None:
        return

    st.caption(
        f"Calculated **{batch.file_count}** report(s) against shared Standard "
        f"**{st.session_state.std_name}**."
    )

    if batch.file_count > 1:
        st.subheader("Per-file summary")
        st.dataframe(
            per_file_summary_table(batch, config),
            use_container_width=True,
            hide_index=True,
        )


def _nav_step(delta: int, total: int) -> None:
    """on_click callback: move report index before widgets are instantiated."""
    current = int(st.session_state.get("report_slider_index", 0))
    new_idx = max(0, min(current + delta, total - 1))
    st.session_state.report_slider_index = new_idx
    st.session_state.report_picker_top = new_idx
    # Reset filters so the newly selected report is not hidden by prior selections.
    _clear_browse_filter_state()


def _nav_sync_picker() -> None:
    """on_change callback: file picker → report index."""
    st.session_state.report_slider_index = int(st.session_state.report_picker_top)
    _clear_browse_filter_state()


def _render_report_nav(batch: BatchResult, idx: int, *, position: str) -> None:
    """
    Prev / file picker / Next controls.

    Top: full picker + buttons. Bottom: buttons only (always reachable after scrolling).
    Navigation state is updated only via on_click / on_change callbacks.
    """
    n = batch.file_count
    names = batch.file_names()
    multi = n > 1

    with st.container(border=True):
        nav_l, nav_c, nav_r = st.columns([1.1, 3.2, 1.1], vertical_alignment="center")
        with nav_l:
            st.button(
                "◀ Previous",
                disabled=idx <= 0 or not multi,
                use_container_width=True,
                key=f"prev_report_{position}",
                on_click=_nav_step,
                args=(-1, n),
            )
        with nav_c:
            if position == "top":
                if not multi:
                    st.markdown(f"**1 / 1** · {names[0]}")
                else:
                    st.selectbox(
                        "Current report",
                        options=list(range(n)),
                        format_func=lambda i: f"{i + 1} / {n}  ·  {names[i]}",
                        key="report_picker_top",
                        label_visibility="collapsed",
                        on_change=_nav_sync_picker,
                    )
            else:
                st.markdown(f"**Report {idx + 1} of {n}** · {names[idx]}")
        with nav_r:
            st.button(
                "Next ▶",
                disabled=idx >= n - 1 or not multi,
                use_container_width=True,
                type="primary",
                key=f"next_report_{position}",
                on_click=_nav_step,
                args=(1, n),
            )

        if multi and position == "top":
            st.progress((idx + 1) / n, text=f"Report {idx + 1} of {n}")


def section_results(config: AppConfig) -> None:
    batch: Optional[BatchResult] = st.session_state.batch_result
    if batch is None or batch.combined is None:
        return

    st.header("F. Results")
    if batch.file_count > 1:
        st.caption(
            f"Browsing **{batch.file_count}** reports. Use **Previous / Next** "
            "(also available at the bottom) or open **All results combined**."
        )
    tab_browse, tab_all = st.tabs(["Browse by report", "All results combined"])

    with tab_browse:
        _section_report_slider(batch, config)

    with tab_all:
        detail = round_for_display(batch.combined.detail, config)
        filtered, _ = _apply_filters(detail, key_prefix="all")
        _render_comparison_and_detail(filtered, config)

        with st.container(border=True):
            st.subheader("Validation issues")
            issues = batch.combined.validation_issues
            if issues is None or issues.empty:
                st.success("No validation issues.")
            else:
                st.dataframe(issues, use_container_width=True, hide_index=True)


def _section_report_slider(batch: BatchResult, config: AppConfig) -> None:
    """Browse individual report results with top and bottom navigation."""
    n = batch.file_count
    if n == 0:
        return

    if st.session_state.report_slider_index >= n:
        st.session_state.report_slider_index = 0
    idx = int(st.session_state.report_slider_index)

    # Sync picker before the selectbox is created (never after).
    if "report_picker_top" not in st.session_state or st.session_state.report_picker_top >= n:
        st.session_state.report_picker_top = idx
    elif int(st.session_state.report_picker_top) != idx:
        # Button callbacks already set both; if only index moved, mirror it here
        # before the widget exists on this run.
        st.session_state.report_picker_top = idx

    fr = batch.file_results[idx]

    _render_report_nav(batch, idx, position="top")

    sheet_note = f" · Worksheet: `{fr.sheet_name}`" if fr.sheet_name else ""
    st.info(
        f"**{fr.name}**{sheet_note}  \n"
        f"{fr.row_count} tasks in this report"
    )

    render_summary_panels(
        fr.result,
        st.session_state.standard_lookup,
        config,
        title_prefix="",
    )

    # One browse filter namespace; cleared on Next/Previous/picker change.
    detail = round_for_display(fr.result.detail, config)
    filtered, _ = _apply_filters(detail, key_prefix="browse")
    _render_comparison_and_detail(filtered, config)

    with st.container(border=True):
        st.subheader("Validation issues")
        issues = fr.result.validation_issues
        if issues is None or issues.empty:
            st.success("No validation issues for this report.")
        else:
            st.dataframe(issues, use_container_width=True, hide_index=True)

    # Bottom navigation — keeps Next available after scrolling long tables
    if n > 1:
        st.divider()
        st.caption("Continue to another report")
        _render_report_nav(batch, idx, position="bottom")


def _build_individual_reports(
    batch: BatchResult,
    lookup,
    config: AppConfig,
    tpl_sheet: Optional[str],
) -> List[Tuple[str, bytes]]:
    """Generate (filename, workbook bytes) for every report in the batch."""
    outputs: List[Tuple[str, bytes]] = []
    used_names: Dict[str, int] = {}
    for fr in batch.file_results:
        report_bytes = generate_report_bytes(
            result=fr.result,
            standard_lookup=lookup,
            config=config,
            template_bytes=st.session_state.tpl_bytes,
            template_sheet=tpl_sheet,
        )
        safe_name = fr.name.rsplit(".", 1)[0]
        base = f"Productivity_Analysis_{safe_name}.xlsx"
        # Avoid colliding names inside a ZIP when two uploads share a filename.
        count = used_names.get(base, 0)
        used_names[base] = count + 1
        if count:
            stem = base.rsplit(".", 1)[0]
            base = f"{stem}_{count + 1}.xlsx"
        outputs.append((base, report_bytes))
    return outputs


def _zip_reports(reports: List[Tuple[str, bytes]]) -> bytes:
    """Pack individual Excel workbooks into a single ZIP archive."""
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        for filename, data in reports:
            zf.writestr(filename, data)
    buffer.seek(0)
    return buffer.read()


def section_export(ctx: Optional[Dict], config: AppConfig) -> None:
    st.header("G. Export")
    batch: Optional[BatchResult] = st.session_state.batch_result
    lookup = st.session_state.standard_lookup
    if batch is None or batch.combined is None or lookup is None:
        st.info("Calculate productivity before exporting.")
        return

    tpl_sheet = ctx.get("tpl_sheet") if ctx else None
    st.caption(
        "Download each Productivity Report as its own Excel workbook, "
        "or download all reports together as a ZIP file. "
        "Each workbook includes Productivity Detail, Summary, and (when applicable) "
        "Validation Issues and Standard Reference sheets."
    )

    try:
        reports = _build_individual_reports(batch, lookup, config, tpl_sheet)
    except Exception as exc:  # noqa: BLE001
        st.error(f"Failed to generate Excel reports: {exc}")
        return

    if len(reports) > 1:
        zip_bytes = _zip_reports(reports)
        st.download_button(
            label=f"Download all ({len(reports)} reports)",
            data=zip_bytes,
            file_name="Productivity_Analysis_All.zip",
            mime="application/zip",
            type="primary",
            key="dl_all_zip",
        )
        st.divider()

    for i, ((filename, report_bytes), fr) in enumerate(
        zip(reports, batch.file_results)
    ):
        with st.container(border=True):
            left, right = st.columns([3, 1.2], vertical_alignment="center")
            with left:
                sheet_note = f" · {fr.sheet_name}" if fr.sheet_name else ""
                st.markdown(
                    f"**{fr.name}**{sheet_note}  \n"
                    f"{fr.row_count} tasks · "
                    f"{fr.result.summary.valid_productivity_tasks} with productivity"
                )
            with right:
                st.download_button(
                    label="Download",
                    data=report_bytes,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"dl_report_{i}_{filename}",
                    use_container_width=True,
                )


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main() -> None:
    _init_session()
    st.title("Productivity Report Analyzer")
    st.write(
        "Compare one or more employee productivity reports against a shared standard "
        "effort table, calculate task-level productivity, browse each report, and "
        "export each report as its own Excel workbook."
    )

    section_upload()
    ctx = section_mapping_and_preview()
    config = section_settings()
    if ctx is not None:
        section_calculate(ctx, config)
    section_results(config)
    section_export(ctx, config)


if __name__ == "__main__":
    main()
