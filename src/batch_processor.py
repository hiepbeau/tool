"""Multi-file batch productivity processing."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Dict, List, Optional, Sequence

import pandas as pd

from .configuration import ISSUE_MISSING_STANDARD, AppConfig
from .productivity_calculator import (
    CalculationResult,
    CalculationSummary,
    calculate_productivity,
    _build_summary,
)
from .validators import StandardLookup


@dataclass
class FileInput:
    """One uploaded Productivity Report ready for calculation."""

    name: str
    dataframe: pd.DataFrame
    sheet_name: Optional[str] = None


@dataclass
class FileResult:
    """Per-file calculation output."""

    name: str
    sheet_name: Optional[str]
    result: CalculationResult
    row_count: int = 0


@dataclass
class BatchResult:
    """Independent per-file results plus a combined aggregate."""

    file_results: List[FileResult] = field(default_factory=list)
    combined: Optional[CalculationResult] = None
    standard_lookup: Optional[StandardLookup] = None
    errors: List[str] = field(default_factory=list)

    @property
    def file_count(self) -> int:
        return len(self.file_results)

    def get_file(self, index: int) -> Optional[FileResult]:
        if 0 <= index < len(self.file_results):
            return self.file_results[index]
        return None

    def file_names(self) -> List[str]:
        return [f.name for f in self.file_results]


def tag_source_file(result: CalculationResult, source_file: str) -> CalculationResult:
    """Return a copy of ``result`` with a Source File column on detail and issues."""
    detail = result.detail.copy()
    if "Source File" in detail.columns:
        detail["Source File"] = source_file
    else:
        detail.insert(0, "Source File", source_file)

    validation = result.validation_issues.copy() if result.validation_issues is not None else pd.DataFrame()
    if not validation.empty:
        if "Source File" in validation.columns:
            validation["Source File"] = source_file
        else:
            validation.insert(0, "Source File", source_file)

    return CalculationResult(
        detail=detail,
        summary=result.summary,
        validation_issues=validation,
    )


def combine_results(file_results: Sequence[FileResult]) -> CalculationResult:
    """
    Concatenate per-file details into one combined result.

    Weighted productivity is recomputed over all valid task pairs
    (Σ standard / Σ actual), not averaged across files.
    """
    if not file_results:
        empty = pd.DataFrame()
        return CalculationResult(
            detail=empty,
            summary=CalculationSummary(),
            validation_issues=empty,
        )

    details = [fr.result.detail.copy() for fr in file_results]
    combined_detail = pd.concat(details, ignore_index=True)

    unmatched_keys = set()
    if not combined_detail.empty and "Issues List" in combined_detail.columns:
        for _, row in combined_detail.iterrows():
            issues = row.get("Issues List") or []
            if ISSUE_MISSING_STANDARD in issues and row.get("Lookup Key"):
                unmatched_keys.add(row["Lookup Key"])

    summary = _build_summary(combined_detail, unmatched_keys)

    validations = [
        fr.result.validation_issues
        for fr in file_results
        if fr.result.validation_issues is not None and not fr.result.validation_issues.empty
    ]
    if validations:
        combined_validation = pd.concat(validations, ignore_index=True)
    else:
        combined_validation = pd.DataFrame()

    return CalculationResult(
        detail=combined_detail,
        summary=summary,
        validation_issues=combined_validation,
    )


def calculate_batch(
    files: Sequence[FileInput],
    productivity_map: Dict[str, str],
    standard_lookup: StandardLookup,
    config: Optional[AppConfig] = None,
) -> BatchResult:
    """
    Calculate productivity independently for each file, then build a combined result.

    Uses one shared Standard lookup for every file. Does not modify core formulas.
    """
    batch = BatchResult(standard_lookup=standard_lookup)
    if not files:
        batch.errors.append("No Productivity Report files were provided.")
        return batch

    cfg = config or AppConfig()
    for file_input in files:
        try:
            raw = calculate_productivity(
                file_input.dataframe,
                productivity_map,
                standard_lookup,
                cfg,
            )
            tagged = tag_source_file(raw, file_input.name)
            batch.file_results.append(
                FileResult(
                    name=file_input.name,
                    sheet_name=file_input.sheet_name,
                    result=tagged,
                    row_count=len(tagged.detail),
                )
            )
        except Exception as exc:  # noqa: BLE001
            batch.errors.append(f"{file_input.name}: {exc}")

    if batch.file_results:
        batch.combined = combine_results(batch.file_results)
    return batch


def per_file_summary_table(batch: BatchResult, config: Optional[AppConfig] = None) -> pd.DataFrame:
    """Build a compact per-file metrics table for UI and export."""
    cfg = config or AppConfig()
    rows = []
    for fr in batch.file_results:
        s = fr.result.summary
        rows.append(
            {
                "Source File": fr.name,
                "Worksheet": fr.sheet_name or "",
                "Total Tasks": s.total_tasks,
                "Valid Tasks": s.valid_productivity_tasks,
                "Missing Standard": s.missing_standard,
                "Ambiguous Standard": s.ambiguous_standard,
                "Invalid Actual Effort": s.invalid_actual_effort + s.zero_actual_effort,
                "Total Actual Effort (MD)": round(s.total_actual_effort, cfg.effort_decimals),
                "Total Standard Effort (MD)": round(s.total_standard_effort, cfg.effort_decimals),
            }
        )
    return pd.DataFrame(rows)


def validate_mapped_columns(
    dataframe: pd.DataFrame,
    column_map: Dict[str, str],
    file_label: str,
) -> List[str]:
    """Ensure every mapped column exists in the dataframe."""
    errors: List[str] = []
    available = set(dataframe.columns)
    for role, header in column_map.items():
        if header not in available:
            errors.append(
                f"{file_label}: mapped column '{header}' (role '{role}') was not found."
            )
    return errors
