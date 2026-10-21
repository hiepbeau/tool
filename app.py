"""
Productivity Report Analyzer – Streamlit UI.

Upload a Productivity Report and a Standard effort table, map columns,
calculate task-level productivity, and download an Excel report.
"""

from __future__ import annotations

from typing import Dict, Optional

import pandas as pd
import streamlit as st

from src.column_mapper import auto_map_columns, validate_mapping
from src.configuration import (
    PRODUCTIVITY_REQUIRED_COLUMNS,
    STANDARD_REQUIRED_COLUMNS,
    AppConfig,
    validate_coefficients,
)
from src.excel_reader import ExcelReadError, inspect_workbook, list_columns, read_table
from src.productivity_calculator import calculate_productivity, round_for_display
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
    "prod_bytes": None,
    "prod_name": None,
    "std_bytes": None,
    "std_name": None,
    "tpl_bytes": None,
    "tpl_name": None,
    "result": None,
    "standard_lookup": None,
    "config_coeffs": None,
}


def _init_session() -> None:
    for key, value in SESSION_DEFAULTS.items():
        if key not in st.session_state:
            st.session_state[key] = value


def _file_bytes(uploaded) -> Optional[bytes]:
    if uploaded is None:
        return None
    return uploaded.getvalue()


# ---------------------------------------------------------------------------
# UI sections
# ---------------------------------------------------------------------------

def section_upload() -> None:
    st.header("A. Upload files")
    st.caption(
        "Upload the Productivity Report and Standard effort table. "
        "An Output Template is optional."
    )

    col1, col2, col3 = st.columns(3)
    with col1:
        prod_file = st.file_uploader(
            "Productivity Report",
            type=["xlsx", "xlsm", "csv"],
            key="prod_uploader",
            help="Employee task report (required each run).",
        )
    with col2:
        std_file = st.file_uploader(
            "Standard (fixed reference)",
            type=["xlsx", "xlsm", "csv"],
            key="std_uploader",
            help="Customer-approved standard effort table. Upload once; it is kept in session.",
        )
    with col3:
        tpl_file = st.file_uploader(
            "Output Template (optional)",
            type=["xlsx", "xlsm"],
            key="tpl_uploader",
        )

    if prod_file is not None:
        st.session_state.prod_bytes = prod_file.getvalue()
        st.session_state.prod_name = prod_file.name
        st.session_state.result = None

    if std_file is not None:
        st.session_state.std_bytes = std_file.getvalue()
        st.session_state.std_name = std_file.name
        st.session_state.result = None

    if tpl_file is not None:
        st.session_state.tpl_bytes = tpl_file.getvalue()
        st.session_state.tpl_name = tpl_file.name

    # Persist Standard across reruns – show status
    if st.session_state.std_bytes and st.session_state.std_name:
        st.success(f"Standard loaded: **{st.session_state.std_name}** (kept in session until replaced).")
    if st.session_state.prod_bytes and st.session_state.prod_name:
        st.info(f"Productivity Report: **{st.session_state.prod_name}**")
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
    if not st.session_state.prod_bytes or not st.session_state.std_bytes:
        st.warning("Upload both a Productivity Report and a Standard file to continue.")
        return None

    try:
        prod_sheet = _sheet_selector(
            "Productivity Report worksheet",
            st.session_state.prod_bytes,
            st.session_state.prod_name,
            "prod_sheet",
        )
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

        prod_df = read_table(
            st.session_state.prod_bytes,
            sheet_name=None if prod_sheet == "CSV" else prod_sheet,
            filename=st.session_state.prod_name,
        )
        std_df = read_table(
            st.session_state.std_bytes,
            sheet_name=None if std_sheet == "CSV" else std_sheet,
            filename=st.session_state.std_name,
        )
    except ExcelReadError as exc:
        st.error(str(exc))
        return None

    prod_cols = list_columns(prod_df)
    std_cols = list_columns(std_df)

    auto_prod = auto_map_columns(prod_cols, PRODUCTIVITY_REQUIRED_COLUMNS)
    auto_std = auto_map_columns(std_cols, STANDARD_REQUIRED_COLUMNS)

    st.subheader("Productivity Report columns")
    prod_map: Dict[str, Optional[str]] = {}
    cols = st.columns(3)
    for i, (role, expected) in enumerate(PRODUCTIVITY_REQUIRED_COLUMNS.items()):
        with cols[i % 3]:
            options = ["— not mapped —"] + prod_cols
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

    if mapping_errors:
        for err in mapping_errors:
            st.error(err)
        mapping_ok = False
    else:
        st.success("All required columns are mapped.")
        mapping_ok = True

    # Section C – preview
    st.header("C. Data preview")
    c1, c2, c3 = st.columns(3)
    c1.metric("Productivity Report rows", len(prod_df))
    c2.metric("Standard rows", len(std_df))
    c3.metric("Productivity columns", len(prod_cols))

    with st.expander("Preview – Productivity Report", expanded=False):
        st.dataframe(prod_df.head(20), use_container_width=True)
    with st.expander("Preview – Standard", expanded=False):
        st.dataframe(std_df.head(20), use_container_width=True)

    return {
        "prod_df": prod_df,
        "std_df": std_df,
        "prod_map": prod_map,
        "std_map": std_map,
        "mapping_ok": mapping_ok,
        "tpl_sheet": tpl_sheet,
    }


def section_settings() -> AppConfig:
    st.header("D. Calculation settings")
    st.markdown(
        """
        <style>
        .formula-highlight-box {
            background: linear-gradient(135deg, #FFF8E7 0%, #FFFCF3 100%);
            border: 1px solid #F0D78C;
            border-left: 5px solid #C69000;
            border-radius: 10px;
            padding: 0.9rem 1.15rem;
            margin: 0.35rem 0 1rem 0;
            box-shadow: 0 1px 2px rgba(16, 24, 40, 0.06);
        }
        .formula-highlight-title {
            font-size: 0.8rem;
            font-weight: 700;
            letter-spacing: 0.04em;
            text-transform: uppercase;
            color: #7A5B00;
            margin-bottom: 0.55rem;
        }
        .formula-highlight-line {
            font-size: 1.02rem;
            font-weight: 600;
            color: #101828;
            font-family: ui-monospace, SFMono-Regular, Menlo, Consolas, monospace;
            margin: 0.25rem 0;
        }
        .formula-highlight-note {
            font-size: 0.85rem;
            color: #667085;
            margin-top: 0.55rem;
        }
        </style>
        <div class="formula-highlight-box">
            <div class="formula-highlight-title">Calculation formulas (fixed)</div>
            <div class="formula-highlight-line">Productivity (%) = Standard Effort ÷ Actual Effort × 100</div>
            <div class="formula-highlight-line">Effort For Specialist (MD) = Coefficient × Actual Effort ÷ 100</div>
        </div>
        """,
        unsafe_allow_html=True,
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
        with st.spinner("Calculating…"):
            try:
                prod_map = {k: v for k, v in ctx["prod_map"].items() if v}
                std_map = {k: v for k, v in ctx["std_map"].items() if v}
                lookup = build_standard_lookup(ctx["std_df"], std_map)
                result = calculate_productivity(
                    ctx["prod_df"], prod_map, lookup, config
                )
                st.session_state.result = result
                st.session_state.standard_lookup = lookup
            except Exception as exc:  # noqa: BLE001
                st.exception(exc)
                return

    result = st.session_state.result
    lookup = st.session_state.standard_lookup
    if result is None:
        return

    s = result.summary
    dup_keys = len(lookup.duplicate_identical_keys)
    amb_keys = len(lookup.ambiguous_keys)
    unmatched = s.unmatched_task_keys
    invalid_effort = s.invalid_actual_effort + s.zero_actual_effort
    actual_md = round(s.total_actual_effort, config.effort_decimals)
    standard_md = round(s.total_standard_effort, config.effort_decimals)

    def _tone(value: int, *, warn_is_bad: bool = True) -> str:
        if value == 0:
            return "ok" if warn_is_bad else "neutral"
        return "warn" if warn_is_bad else "neutral"

    def _metric_card(label: str, value, tone: str = "neutral", unit: str = "") -> str:
        unit_html = f'<span class="sum-unit">{unit}</span>' if unit else ""
        return (
            f'<div class="sum-card tone-{tone}">'
            f'<div class="sum-label">{label}</div>'
            f'<div class="sum-value">{value}{unit_html}</div>'
            f"</div>"
        )

    st.markdown(
        f"""
        <style>
        .sum-section {{
            margin: 0.25rem 0 1.15rem 0;
        }}
        .sum-section-title {{
            font-size: 1.05rem;
            font-weight: 700;
            color: #101828;
            margin: 0 0 0.65rem 0;
        }}
        .sum-section-subtitle {{
            font-size: 0.78rem;
            font-weight: 700;
            letter-spacing: 0.04em;
            text-transform: uppercase;
            color: #667085;
            margin: 0.85rem 0 0.45rem 0;
        }}
        .sum-grid {{
            display: grid;
            gap: 0.75rem;
        }}
        .sum-grid.cols-3 {{ grid-template-columns: repeat(3, minmax(0, 1fr)); }}
        .sum-grid.cols-2 {{ grid-template-columns: repeat(2, minmax(0, 1fr)); }}
        .sum-card {{
            background: #FFFFFF;
            border: 1px solid #E4E7EC;
            border-radius: 10px;
            padding: 0.85rem 1rem;
            box-shadow: 0 1px 2px rgba(16, 24, 40, 0.04);
            border-top: 3px solid #D0D5DD;
        }}
        .sum-card.tone-neutral {{ border-top-color: #2F5496; }}
        .sum-card.tone-ok {{
            border-top-color: #2E7D4F;
            background: linear-gradient(180deg, #F6FBF8 0%, #FFFFFF 70%);
        }}
        .sum-card.tone-warn {{
            border-top-color: #C69000;
            background: linear-gradient(180deg, #FFFBF0 0%, #FFFFFF 70%);
        }}
        .sum-card.tone-accent-blue {{
            border-top-color: #2F5496;
            background: linear-gradient(135deg, #E8F1FB 0%, #F8FBFE 100%);
            border-color: #B6D0EE;
        }}
        .sum-card.tone-accent-green {{
            border-top-color: #2E7D4F;
            background: linear-gradient(135deg, #E7F6ED 0%, #F7FCF9 100%);
            border-color: #B7E0C5;
        }}
        .sum-label {{
            font-size: 0.82rem;
            font-weight: 600;
            color: #475467;
            margin-bottom: 0.35rem;
            line-height: 1.3;
            min-height: 2.2em;
        }}
        .sum-value {{
            font-size: 1.55rem;
            font-weight: 700;
            color: #101828;
            line-height: 1.15;
        }}
        .sum-unit {{
            font-size: 0.85rem;
            font-weight: 600;
            color: #667085;
            margin-left: 0.3rem;
        }}
        @media (max-width: 900px) {{
            .sum-grid.cols-3, .sum-grid.cols-2 {{
                grid-template-columns: 1fr;
            }}
        }}
        </style>

        <div class="sum-section">
            <div class="sum-section-title">Match overview</div>
            <div class="sum-grid cols-3">
                {_metric_card("Duplicate identical standard keys", dup_keys, _tone(dup_keys))}
                {_metric_card("Ambiguous standard keys", amb_keys, _tone(amb_keys))}
                {_metric_card("Unmatched task keys", unmatched, _tone(unmatched))}
            </div>
        </div>

        <div class="sum-section">
            <div class="sum-section-title">Calculation summary</div>
            <div class="sum-section-subtitle">Coverage</div>
            <div class="sum-grid cols-2">
                {_metric_card("Total source tasks", s.total_tasks, "neutral")}
                {_metric_card("Valid productivity tasks", s.valid_productivity_tasks, "neutral")}
            </div>
            <div class="sum-section-subtitle">Validation flags</div>
            <div class="sum-grid cols-3">
                {_metric_card("Missing standard", s.missing_standard, _tone(s.missing_standard))}
                {_metric_card("Ambiguous standard", s.ambiguous_standard, _tone(s.ambiguous_standard))}
                {_metric_card("Invalid actual effort", invalid_effort, _tone(invalid_effort))}
            </div>
            <div class="sum-section-subtitle">Effort totals</div>
            <div class="sum-grid cols-2">
                {_metric_card("Total actual effort (valid)", actual_md, "accent-blue", "MD")}
                {_metric_card("Total standard effort (valid)", standard_md, "accent-green", "MD")}
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def section_results(config: AppConfig) -> None:
    result = st.session_state.result
    if result is None:
        return

    st.header("F. Results")
    detail = round_for_display(result.detail, config)

    st.subheader("Filters")
    f1, f2, f3 = st.columns(3)
    employees = sorted({e for e in detail["Employee Name"].dropna().unique()})
    complexities = sorted({c for c in detail["Complexity"].dropna().unique()})
    task_types = sorted({t for t in detail["Task Type"].dropna().unique()})
    changes = sorted({c for c in detail["Change/New"].dropna().unique()})
    levels = sorted({lv for lv in detail["Professional Level"].dropna().unique()})
    statuses = sorted(detail["Validation Status"].dropna().unique())

    sel_emp = f1.multiselect("Employee Name", employees)
    sel_cx = f2.multiselect("Complexity", complexities)
    sel_tt = f3.multiselect("Task Type", task_types)

    f4, f5, f6 = st.columns(3)
    sel_ch = f4.multiselect("Change/New", changes)
    sel_lv = f5.multiselect("Professional Level", levels)
    sel_st = f6.multiselect("Validation Status", statuses)

    date_min = detail["Finished Date"].min()
    date_max = detail["Finished Date"].max()
    date_range = None
    if pd.notna(date_min) and pd.notna(date_max):
        try:
            d1, d2 = st.date_input(
                "Finished Date range",
                value=(pd.Timestamp(date_min).date(), pd.Timestamp(date_max).date()),
            )
            date_range = (pd.Timestamp(d1), pd.Timestamp(d2) + pd.Timedelta(days=1))
        except Exception:  # noqa: BLE001
            date_range = None

    filtered = detail
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
    if date_range is not None:
        dates = pd.to_datetime(filtered["Finished Date"], errors="coerce")
        filtered = filtered[(dates >= date_range[0]) & (dates < date_range[1])]

    # --- Productivity Report comparison (above calculated detail) ---
    st.subheader("Productivity Report")
    st.caption(
        "Side-by-side view for comparison. "
        "EffortManday is highlighted in yellow when it differs from Standard Effort, "
        "and in green when it matches."
    )
    compare_cols = [
        c
        for c in [
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

    if compare.empty:
        st.info("No rows match the current filters.")
    else:
        effort_fmt = f"{{:.{config.effort_decimals}f}}"
        formatters = {}
        for col in ("EffortManday", "Standard Effort"):
            if col in compare.columns:
                formatters[col] = effort_fmt
        styled = (
            compare.style.apply(_highlight_effort_diff, axis=1)
            .format(formatters, na_rep="")
        )
        st.dataframe(styled, use_container_width=True, hide_index=True)

    # --- Calculated detail ---
    st.subheader("Calculated results")
    has_validation_issues = (
        "Validation Status" in filtered.columns
        and (filtered["Validation Status"] != "OK").any()
    )
    show_cols = [
        c
        for c in [
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
    st.dataframe(display, use_container_width=True, hide_index=True)

    st.subheader("Validation issues")
    if result.validation_issues is None or result.validation_issues.empty:
        st.success("No validation issues.")
    else:
        st.dataframe(result.validation_issues, use_container_width=True, hide_index=True)


def section_export(ctx: Optional[Dict], config: AppConfig) -> None:
    st.header("G. Export")
    result = st.session_state.result
    lookup = st.session_state.standard_lookup
    if result is None or lookup is None:
        st.info("Calculate productivity before exporting.")
        return

    tpl_sheet = ctx.get("tpl_sheet") if ctx else None
    try:
        report_bytes = generate_report_bytes(
            result=result,
            standard_lookup=lookup,
            config=config,
            template_bytes=st.session_state.tpl_bytes,
            template_sheet=tpl_sheet,
        )
    except Exception as exc:  # noqa: BLE001
        st.error(f"Failed to generate Excel report: {exc}")
        return

    st.download_button(
        label="Download Excel Report",
        data=report_bytes,
        file_name="Productivity_Analysis_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )
    st.caption(
        "The workbook includes Productivity Detail, Summary, and (when applicable) "
        "Validation Issues and Standard Reference sheets."
    )


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main() -> None:
    _init_session()
    st.title("Productivity Report Analyzer")
    st.write(
        "Compare employee productivity tasks against a standard effort table, "
        "calculate task-level productivity, and export a detail report to Excel."
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
