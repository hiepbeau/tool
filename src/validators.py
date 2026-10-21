"""Validation helpers for standard tables and calculated results."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd

from .column_mapper import (
    build_lookup_key,
    normalize_change_or_new,
    normalize_difficulty,
    normalize_work_category,
    parse_standard_effort,
)
from .configuration import (
    ISSUE_AMBIGUOUS_STANDARD,
    ISSUE_INVALID_STANDARD_EFFORT,
)


@dataclass
class StandardLookup:
    """Normalized standard effort lookup with duplicate/conflict tracking."""

    # key → effort (only unambiguous keys)
    efforts: Dict[str, float] = field(default_factory=dict)
    # key → list of efforts seen (for audit)
    all_efforts: Dict[str, List[float]] = field(default_factory=dict)
    # keys with conflicting effort values
    ambiguous_keys: Dict[str, List[float]] = field(default_factory=dict)
    # keys consolidated from identical duplicates
    duplicate_identical_keys: Dict[str, float] = field(default_factory=dict)
    # rows that failed validation
    invalid_rows: List[Dict[str, Any]] = field(default_factory=list)
    # full normalized table for audit worksheet
    normalized_rows: List[Dict[str, Any]] = field(default_factory=list)

    def get(self, key: Optional[str]) -> Tuple[Optional[float], Optional[str]]:
        """
        Return (effort, issue_code).

        issue_code is AMBIGUOUS_STANDARD, MISSING_STANDARD, or None.
        """
        if key is None:
            return None, "MISSING_STANDARD"
        if key in self.ambiguous_keys:
            return None, ISSUE_AMBIGUOUS_STANDARD
        if key in self.efforts:
            return self.efforts[key], None
        return None, "MISSING_STANDARD"


def build_standard_lookup(
    standard_df: pd.DataFrame,
    column_map: Dict[str, str],
) -> StandardLookup:
    """
    Build a standard effort lookup from the Standard worksheet.

    Rules:
    - Lookup key = Work Category + New/Change + Difficulty (normalized).
    - Identical duplicate efforts are consolidated and reported.
    - Conflicting efforts mark the key as ambiguous.
    - Zero / negative / invalid standard efforts are rejected.
    """
    lookup = StandardLookup()
    cat_col = column_map["work_category"]
    nc_col = column_map["new_change"]
    diff_col = column_map["difficulty"]
    effort_col = column_map["standard_effort"]

    grouped: Dict[str, List[Tuple[float, Dict[str, Any]]]] = {}

    for _, row in standard_df.iterrows():
        source_row = int(row["_source_row"]) if "_source_row" in row.index else None
        raw_cat = row.get(cat_col)
        raw_nc = row.get(nc_col)
        raw_diff = row.get(diff_col)
        raw_effort = row.get(effort_col)

        # Skip completely empty trailing rows.
        if all(
            (v is None or (isinstance(v, float) and pd.isna(v)) or str(v).strip() == "")
            for v in (raw_cat, raw_nc, raw_diff, raw_effort)
        ):
            continue

        cat = normalize_work_category(raw_cat)
        nc = normalize_change_or_new(raw_nc)
        diff = normalize_difficulty(raw_diff)
        effort, effort_err = parse_standard_effort(raw_effort)

        record = {
            "Source Row": source_row,
            "Work Category (raw)": raw_cat,
            "New/Change (raw)": raw_nc,
            "Difficulty (raw)": raw_diff,
            "Standard Effort (raw)": raw_effort,
            "Work Category": cat,
            "New/Change": nc,
            "Difficulty": diff,
            "Standard Effort (MD)": effort,
            "Lookup Key": None,
            "Status": "OK",
            "Issues": "",
        }

        issues: List[str] = []
        if not cat or not nc or not diff:
            issues.append("MISSING_CLASSIFICATION")
        if effort_err:
            issues.append(ISSUE_INVALID_STANDARD_EFFORT)
            issues.append(effort_err)

        key = build_lookup_key(cat, nc, diff)
        record["Lookup Key"] = key

        if issues or effort is None or key is None:
            record["Status"] = "INVALID"
            record["Issues"] = "; ".join(issues)
            lookup.invalid_rows.append(record)
            lookup.normalized_rows.append(record)
            continue

        grouped.setdefault(key, []).append((effort, record))
        lookup.normalized_rows.append(record)

    for key, entries in grouped.items():
        efforts = [e for e, _ in entries]
        unique = sorted(set(efforts))
        lookup.all_efforts[key] = efforts

        if len(unique) == 1:
            lookup.efforts[key] = unique[0]
            if len(efforts) > 1:
                lookup.duplicate_identical_keys[key] = unique[0]
                for _, rec in entries:
                    rec["Status"] = "DUPLICATE_IDENTICAL"
                    rec["Issues"] = f"Identical duplicate consolidated to {unique[0]}"
        else:
            lookup.ambiguous_keys[key] = efforts
            for _, rec in entries:
                rec["Status"] = "AMBIGUOUS"
                rec["Issues"] = f"Conflicting efforts: {efforts}"

    return lookup


def summarize_issues(issues: List[str]) -> str:
    """Join issue codes into a stable semicolon-separated string."""
    # Preserve order, drop empties, keep uniqueness.
    seen = set()
    ordered: List[str] = []
    for item in issues:
        if not item or item in seen:
            continue
        seen.add(item)
        ordered.append(item)
    return "; ".join(ordered)
