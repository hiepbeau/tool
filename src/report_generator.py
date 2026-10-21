"""Excel report generation and template-aware export."""

from __future__ import annotations

import io
from datetime import datetime
from typing import Any, Dict, List, Optional

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows

from .configuration import OUTPUT_AUDIT_COLUMNS, OUTPUT_DETAIL_COLUMNS, AppConfig
from .productivity_calculator import CalculationResult, round_for_display
from .validators import StandardLookup

# Columns omitted from every exported worksheet.
EXPORT_EXCLUDED_COLUMNS = frozenset(
    {"Source Row", "Employee Coefficient", "Validation Status"}
)

# Maps template header → detail column name.
TEMPLATE_COLUMN_MAP: Dict[str, str] = {
    "Employee Name": "Employee Name",
    "Title": "Professional Level",
    "Finished Date": "Finished Date",
    "Complexity": "Complexity",
    "Total Effort (MD)": "Total Effort (MD)",
    "Standard Effort": "Standard Effort",
    "Productivity": "Productivity",
    "Effort For Specialist": "Effort For Specialist",
}


def generate_report_bytes(
    result: CalculationResult,
    standard_lookup: StandardLookup,
    config: Optional[AppConfig] = None,
    template_bytes: Optional[bytes] = None,
    template_sheet: Optional[str] = None,
) -> bytes:
    """
    Build an Excel workbook and return it as bytes.

    If ``template_bytes`` is provided, attempt to write calculated values into
    the template's column layout while preserving formatting. Unmapped template
    columns are left blank (sample rows are cleared). Additional audit sheets
    are appended.
    """
    cfg = config or AppConfig()
    detail = round_for_display(result.detail, cfg)

    if template_bytes:
        try:
            return _export_with_template(
                detail=detail,
                result=result,
                standard_lookup=standard_lookup,
                config=cfg,
                template_bytes=template_bytes,
                template_sheet=template_sheet,
            )
        except TemplateExportError:
            # Fall back to the standard formatted report.
            pass

    return _export_default(detail, result, standard_lookup, cfg)


class TemplateExportError(Exception):
    """Raised when a template cannot be safely used."""


def _export_with_template(
    detail: pd.DataFrame,
    result: CalculationResult,
    standard_lookup: StandardLookup,
    config: AppConfig,
    template_bytes: bytes,
    template_sheet: Optional[str],
) -> bytes:
    wb = load_workbook(io.BytesIO(template_bytes))
    sheet_name = template_sheet or wb.sheetnames[0]
    if sheet_name not in wb.sheetnames:
        raise TemplateExportError(f"Template sheet '{sheet_name}' not found.")

    ws = wb[sheet_name]
    headers = _read_header_row(ws)
    if not headers:
        raise TemplateExportError("Template has no header row.")

    # Require at least the core calculated columns.
    required = {
        "Employee Name",
        "Finished Date",
        "Complexity",
        "Total Effort (MD)",
        "Standard Effort",
        "Productivity",
        "Effort For Specialist",
    }
    present = set(headers.values())
    if not required.issubset(present):
        missing = required - present
        raise TemplateExportError(
            f"Template is missing required columns: {sorted(missing)}"
        )

    # Clear existing data rows (keep header + formatting on row 1).
    if ws.max_row > 1:
        ws.delete_rows(2, ws.max_row - 1)

    # Capture number formats from a hypothetical second row if present in template
    # styles – use header row adjacent styles / reasonable defaults.
    num_formats = _infer_number_formats(headers)

    for r_idx in range(len(detail)):
        excel_row = r_idx + 2
        record = detail.iloc[r_idx]
        for col_idx, header in headers.items():
            source_col = TEMPLATE_COLUMN_MAP.get(header)
            value: Any = None
            if source_col and source_col in record.index:
                value = record[source_col]
                if pd.isna(value):
                    value = None
            cell = ws.cell(
                row=excel_row,
                column=col_idx,
                value=_excel_value(value, header, config),
            )
            if header in num_formats and isinstance(cell.value, (int, float)):
                cell.number_format = num_formats[header]

    # Rename primary sheet for clarity when practical.
    if sheet_name == "Sheet1":
        ws.title = "Productivity Detail"

    _append_support_sheets(wb, detail, result, standard_lookup, config)
    return _workbook_to_bytes(wb)


def _export_default(
    detail: pd.DataFrame,
    result: CalculationResult,
    standard_lookup: StandardLookup,
    config: AppConfig,
) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Productivity Detail"

    export_cols = _default_detail_columns(detail)
    _write_dataframe(ws, detail[export_cols], config)
    _append_support_sheets(wb, detail, result, standard_lookup, config)
    return _workbook_to_bytes(wb)


def _default_detail_columns(detail: pd.DataFrame) -> List[str]:
    preferred = OUTPUT_DETAIL_COLUMNS + [
        c for c in OUTPUT_AUDIT_COLUMNS if c not in OUTPUT_DETAIL_COLUMNS
    ]
    return [
        c
        for c in preferred
        if c in detail.columns and c not in EXPORT_EXCLUDED_COLUMNS
    ]


def _drop_excluded_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Remove columns that must not appear in Excel exports."""
    keep = [c for c in df.columns if c not in EXPORT_EXCLUDED_COLUMNS]
    return df[keep]


def _append_support_sheets(
    wb: Workbook,
    detail: pd.DataFrame,
    result: CalculationResult,
    standard_lookup: StandardLookup,
    config: AppConfig,
) -> None:
    # Summary
    ws_sum = wb.create_sheet("Summary")
    summary_rows = _summary_rows(result, config)
    for r_idx, (label, value) in enumerate(summary_rows, start=1):
        ws_sum.cell(row=r_idx, column=1, value=label).font = Font(bold=True)
        ws_sum.cell(row=r_idx, column=2, value=value)
    ws_sum.column_dimensions["A"].width = 40
    ws_sum.column_dimensions["B"].width = 24

    # Validation Issues – only if there are issues
    if result.validation_issues is not None and not result.validation_issues.empty:
        ws_val = wb.create_sheet("Validation Issues")
        _write_dataframe(ws_val, _drop_excluded_columns(result.validation_issues), config)

    # Standard Reference
    if standard_lookup.normalized_rows:
        std_df = _drop_excluded_columns(pd.DataFrame(standard_lookup.normalized_rows))
        ws_std = wb.create_sheet("Standard Reference")
        _write_dataframe(ws_std, std_df, config)

        if standard_lookup.ambiguous_keys or standard_lookup.duplicate_identical_keys:
            audit_rows = []
            for key, efforts in standard_lookup.ambiguous_keys.items():
                audit_rows.append(
                    {
                        "Lookup Key": key,
                        "Type": "AMBIGUOUS",
                        "Efforts": ", ".join(str(e) for e in efforts),
                    }
                )
            for key, effort in standard_lookup.duplicate_identical_keys.items():
                audit_rows.append(
                    {
                        "Lookup Key": key,
                        "Type": "DUPLICATE_IDENTICAL",
                        "Efforts": str(effort),
                    }
                )
            if audit_rows:
                ws_dup = wb.create_sheet("Standard Duplicates")
                _write_dataframe(ws_dup, pd.DataFrame(audit_rows), config)


def _summary_rows(result: CalculationResult, config: AppConfig) -> List[tuple]:
    s = result.summary
    return [
        ("Total source tasks", s.total_tasks),
        ("Valid tasks with productivity", s.valid_productivity_tasks),
        ("Tasks missing a standard", s.missing_standard),
        ("Tasks with ambiguous standards", s.ambiguous_standard),
        ("Tasks with invalid actual effort", s.invalid_actual_effort),
        ("Tasks with zero actual effort", s.zero_actual_effort),
        ("Total actual effort (valid)", round(s.total_actual_effort, config.effort_decimals)),
        ("Total standard effort (valid)", round(s.total_standard_effort, config.effort_decimals)),
        ("Generated at", datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
    ]


def _write_dataframe(ws, df: pd.DataFrame, config: AppConfig) -> None:
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="2F5496")
    thin = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )

    for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), start=1):
        for c_idx, value in enumerate(row, start=1):
            header = df.columns[c_idx - 1] if r_idx > 1 else None
            if header == "Productivity":
                cell_value = _format_productivity_percent(value, config.productivity_decimals)
            else:
                cell_value = _clean_cell(value)
            cell = ws.cell(row=r_idx, column=c_idx, value=cell_value)
            cell.border = thin
            if r_idx == 1:
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = Alignment(horizontal="center", wrap_text=True)
            else:
                assert header is not None
                if header in ("Total Effort (MD)", "Standard Effort", "Effort For Specialist"):
                    if isinstance(cell.value, (int, float)):
                        cell.number_format = f"0.{'0' * config.effort_decimals}"
                elif header == "Finished Date" and isinstance(cell.value, datetime):
                    cell.number_format = "YYYY-MM-DD"

    for idx, col in enumerate(df.columns, start=1):
        width = min(max(12, len(str(col)) + 2), 40)
        ws.column_dimensions[get_column_letter(idx)].width = width
    ws.auto_filter.ref = ws.dimensions
    ws.freeze_panes = "A2"


def _clean_cell(value: Any) -> Any:
    if isinstance(value, list):
        return "; ".join(str(v) for v in value)
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, pd.Timestamp):
        return value.to_pydatetime()
    return value


def _format_productivity_percent(value: Any, decimals: int) -> Any:
    """Format productivity as a percentage string for Excel export (e.g. 200.00%)."""
    value = _clean_cell(value)
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return f"{float(value):.{decimals}f}%"
    return value


def _excel_value(value: Any, header: str, config: Optional[AppConfig] = None) -> Any:
    cfg = config or AppConfig()
    if header == "Productivity":
        return _format_productivity_percent(value, cfg.productivity_decimals)
    value = _clean_cell(value)
    if header == "Finished Date" and value is not None:
        # Template sample uses YYYYMM integers; keep datetime when available
        # so Excel number formats remain useful. Callers may override.
        if isinstance(value, datetime):
            return value
        if isinstance(value, pd.Timestamp):
            return value.to_pydatetime()
    return value


def _read_header_row(ws) -> Dict[int, str]:
    headers: Dict[int, str] = {}
    for cell in ws[1]:
        if cell.value is not None and str(cell.value).strip():
            headers[cell.column] = str(cell.value).strip()
    return headers


def _infer_number_formats(headers: Dict[int, str]) -> Dict[str, str]:
    formats: Dict[str, str] = {}
    for header in headers.values():
        if header in ("Total Effort (MD)", "Standard Effort", "Effort For Specialist"):
            formats[header] = "0.00"
        elif header == "Finished Date":
            formats[header] = "YYYY-MM-DD"
    return formats


def _workbook_to_bytes(wb: Workbook) -> bytes:
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer.read()


def preview_columns_from_template(template_bytes: bytes, sheet_name: Optional[str] = None) -> List[str]:
    """Return header names from an output template worksheet."""
    wb = load_workbook(io.BytesIO(template_bytes), read_only=True, data_only=True)
    name = sheet_name or wb.sheetnames[0]
    ws = wb[name]
    headers = [str(c.value).strip() for c in next(ws.iter_rows(min_row=1, max_row=1)) if c.value]
    wb.close()
    return headers
