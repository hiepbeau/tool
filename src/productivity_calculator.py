"""Productivity calculation and summary metrics."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional

import pandas as pd

from .column_mapper import (
    build_lookup_key,
    normalize_change_or_new,
    normalize_complexity,
    normalize_task_type,
    parse_assignee,
    parse_effort_manday,
    trim_text,
)
from .configuration import (
    AppConfig,
    ISSUE_AMBIGUOUS_STANDARD,
    ISSUE_BLANK_EMPLOYEE,
    ISSUE_INVALID_ACTUAL_EFFORT,
    ISSUE_MISSING_CLASSIFICATION,
    ISSUE_MISSING_STANDARD,
    ISSUE_UNKNOWN_LEVEL,
    ISSUE_ZERO_ACTUAL_EFFORT,
)
from .validators import StandardLookup, summarize_issues


@dataclass
class CalculationSummary:
    """Aggregate metrics after a productivity run."""

    total_tasks: int = 0
    valid_productivity_tasks: int = 0
    missing_standard: int = 0
    ambiguous_standard: int = 0
    invalid_actual_effort: int = 0
    zero_actual_effort: int = 0
    unknown_professional_level: int = 0
    blank_employee_name: int = 0
    total_actual_effort: float = 0.0
    total_standard_effort: float = 0.0
    weighted_productivity: Optional[float] = None
    excluded_from_weighted: int = 0
    unmatched_task_keys: int = 0
    detected_levels: Dict[str, int] = field(default_factory=dict)


@dataclass
class CalculationResult:
    """Full calculation output."""

    detail: pd.DataFrame
    summary: CalculationSummary
    validation_issues: pd.DataFrame


def calculate_productivity(
    productivity_df: pd.DataFrame,
    productivity_map: Dict[str, str],
    standard_lookup: StandardLookup,
    config: Optional[AppConfig] = None,
) -> CalculationResult:
    """
    Left-join each Productivity Report task to the Standard lookup and
    compute productivity metrics.

    Every source task is preserved. Invalid / unmatched tasks are flagged
    rather than dropped or silently zero-filled.
    """
    cfg = config or AppConfig()
    rows: List[Dict[str, Any]] = []

    task_col = productivity_map["task_type"]
    change_col = productivity_map["change_or_new"]
    complexity_col = productivity_map["complexity"]
    effort_col = productivity_map["effort_manday"]
    assignee_col = productivity_map["assignee"]
    date_col = productivity_map["finished_date"]

    unmatched_keys = set()

    for _, src in productivity_df.iterrows():
        source_row = int(src["_source_row"]) if "_source_row" in src.index else None
        issues: List[str] = []

        raw_task = src.get(task_col)
        raw_change = src.get(change_col)
        raw_complexity = src.get(complexity_col)
        raw_effort = src.get(effort_col)
        raw_assignee = src.get(assignee_col)
        raw_date = src.get(date_col)

        # Optional identity columns when present.
        task_key = trim_text(src.get("Key")) if "Key" in src.index else None

        task_type = normalize_task_type(raw_task)
        change_or_new = normalize_change_or_new(raw_change)
        complexity = normalize_complexity(raw_complexity)

        if not task_type or not change_or_new or not complexity:
            issues.append(ISSUE_MISSING_CLASSIFICATION)

        actual_effort, effort_err = parse_effort_manday(raw_effort)
        if effort_err:
            issues.append(ISSUE_INVALID_ACTUAL_EFFORT)
        elif actual_effort == 0:
            issues.append(ISSUE_ZERO_ACTUAL_EFFORT)

        assignee_info = parse_assignee(raw_assignee, cfg.coefficients)
        for issue in assignee_info["issues"]:
            if issue == "UNKNOWN_PROFESSIONAL_LEVEL":
                issues.append(ISSUE_UNKNOWN_LEVEL)
            elif issue == "BLANK_EMPLOYEE_NAME":
                issues.append(ISSUE_BLANK_EMPLOYEE)
            elif issue not in issues:
                issues.append(issue)

        lookup_key = build_lookup_key(task_type, change_or_new, complexity)
        standard_effort: Optional[float] = None
        if lookup_key is None:
            if ISSUE_MISSING_CLASSIFICATION not in issues:
                issues.append(ISSUE_MISSING_CLASSIFICATION)
        else:
            standard_effort, std_issue = standard_lookup.get(lookup_key)
            if std_issue == ISSUE_AMBIGUOUS_STANDARD:
                issues.append(ISSUE_AMBIGUOUS_STANDARD)
            elif std_issue == "MISSING_STANDARD" or std_issue == ISSUE_MISSING_STANDARD:
                issues.append(ISSUE_MISSING_STANDARD)
                unmatched_keys.add(lookup_key)

        productivity: Optional[float] = None
        can_calc_productivity = (
            actual_effort is not None
            and actual_effort > 0
            and standard_effort is not None
            and standard_effort > 0
        )
        if can_calc_productivity:
            # Full precision internally; round only for display/export.
            productivity = (standard_effort / actual_effort) * 100.0

        effort_for_specialist: Optional[float] = None
        coefficient = assignee_info["coefficient"]
        if actual_effort is not None and coefficient is not None:
            effort_for_specialist = coefficient * actual_effort / 100.0

        status = "OK" if not issues else "ISSUES"
        # Productivity is still valid even with unknown level (per Test 8).
        productivity_valid = can_calc_productivity

        finished_date = _coerce_date(raw_date)

        rows.append(
            {
                "Source Row": source_row,
                "Task Key": task_key,
                "Employee Name": assignee_info["employee_name"],
                "Professional Level": assignee_info["professional_level"],
                "Employee Coefficient": coefficient,
                "Finished Date": finished_date,
                "Task Type": task_type,
                "Change/New": change_or_new,
                "Complexity": complexity,
                "Total Effort (MD)": actual_effort,
                "Standard Effort": standard_effort,
                "Productivity": productivity,
                "Effort For Specialist": effort_for_specialist,
                "Lookup Key": lookup_key,
                "Assignee (raw)": assignee_info["original_assignee"],
                "TaskType (raw)": trim_text(raw_task),
                "ChangeOrNew (raw)": trim_text(raw_change),
                "Complexity (raw)": trim_text(raw_complexity),
                "EffortManday (raw)": trim_text(raw_effort)
                if not isinstance(raw_effort, (int, float))
                else raw_effort,
                "Validation Status": status,
                "Validation Issues": summarize_issues(issues),
                "Issues List": list(issues),
                "Productivity Valid": productivity_valid,
            }
        )

    detail = pd.DataFrame(rows)
    summary = _build_summary(detail, unmatched_keys)
    validation = detail[detail["Validation Status"] != "OK"].copy()
    if not validation.empty:
        validation = validation[
            [
                "Source Row",
                "Task Key",
                "Employee Name",
                "Task Type",
                "Change/New",
                "Complexity",
                "Total Effort (MD)",
                "Standard Effort",
                "Validation Issues",
            ]
        ].reset_index(drop=True)

    return CalculationResult(detail=detail, summary=summary, validation_issues=validation)


def _coerce_date(value: Any) -> Any:
    """Preserve dates; attempt to parse strings; leave invalid as original."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, pd.Timestamp):
        return value.to_pydatetime()
    try:
        parsed = pd.to_datetime(value, errors="coerce")
        if pd.isna(parsed):
            return value
        return parsed.to_pydatetime() if hasattr(parsed, "to_pydatetime") else parsed
    except Exception:  # noqa: BLE001
        return value


def _build_summary(detail: pd.DataFrame, unmatched_keys: set) -> CalculationSummary:
    summary = CalculationSummary()
    summary.total_tasks = len(detail)
    if detail.empty:
        summary.weighted_productivity = None
        return summary

    issues_series = detail["Issues List"]

    def _count(code: str) -> int:
        return int(issues_series.apply(lambda xs: code in xs).sum())

    summary.missing_standard = _count(ISSUE_MISSING_STANDARD)
    summary.ambiguous_standard = _count(ISSUE_AMBIGUOUS_STANDARD)
    summary.invalid_actual_effort = _count(ISSUE_INVALID_ACTUAL_EFFORT)
    summary.zero_actual_effort = _count(ISSUE_ZERO_ACTUAL_EFFORT)
    summary.unknown_professional_level = _count(ISSUE_UNKNOWN_LEVEL)
    summary.blank_employee_name = _count(ISSUE_BLANK_EMPLOYEE)
    summary.unmatched_task_keys = len(unmatched_keys)

    valid_mask = detail["Productivity Valid"].astype(bool)
    summary.valid_productivity_tasks = int(valid_mask.sum())
    summary.excluded_from_weighted = int((~valid_mask).sum())

    if summary.valid_productivity_tasks > 0:
        valid = detail.loc[valid_mask]
        summary.total_actual_effort = float(valid["Total Effort (MD)"].sum())
        summary.total_standard_effort = float(valid["Standard Effort"].sum())
        if summary.total_actual_effort > 0:
            summary.weighted_productivity = (
                summary.total_standard_effort / summary.total_actual_effort
            ) * 100.0
        else:
            summary.weighted_productivity = None
    else:
        summary.weighted_productivity = None
        # Still report totals over tasks that have numeric actuals for transparency.
        actuals = detail["Total Effort (MD)"].dropna()
        summary.total_actual_effort = float(actuals.sum()) if not actuals.empty else 0.0

    levels = detail["Professional Level"].dropna()
    summary.detected_levels = {
        str(k): int(v) for k, v in levels.value_counts().to_dict().items()
    }
    return summary


def round_for_display(
    detail: pd.DataFrame,
    config: Optional[AppConfig] = None,
) -> pd.DataFrame:
    """Return a copy with effort/productivity rounded for UI and export."""
    cfg = config or AppConfig()
    out = detail.copy()
    for col in ("Total Effort (MD)", "Standard Effort", "Effort For Specialist"):
        if col in out.columns:
            out[col] = out[col].apply(
                lambda v: round(v, cfg.effort_decimals) if isinstance(v, (int, float)) else v
            )
    if "Productivity" in out.columns:
        out["Productivity"] = out["Productivity"].apply(
            lambda v: round(v, cfg.productivity_decimals)
            if isinstance(v, (int, float))
            else v
        )
    return out
