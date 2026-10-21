"""Column mapping, value normalization, and assignee parsing."""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd

from .configuration import (
    CHANGE_OR_NEW_ALIASES,
    COMPLEXITY_PREFIX,
    DEFAULT_COEFFICIENTS,
    EFFORT_PREFIX,
    PRODUCTIVITY_REQUIRED_COLUMNS,
    STANDARD_REQUIRED_COLUMNS,
    TASK_TYPE_PREFIX,
)


@dataclass
class ColumnMapping:
    """Maps logical field roles to actual Excel header names."""

    productivity: Dict[str, str] = field(
        default_factory=lambda: dict(PRODUCTIVITY_REQUIRED_COLUMNS)
    )
    standard: Dict[str, str] = field(
        default_factory=lambda: dict(STANDARD_REQUIRED_COLUMNS)
    )


def _normalize_header(name: str) -> str:
    return re.sub(r"\s+", " ", str(name).strip().lower())


def auto_map_columns(
    available: List[str],
    required: Dict[str, str],
) -> Dict[str, Optional[str]]:
    """
    Suggest a mapping from logical roles to available headers.

    Exact (case-insensitive, whitespace-normalized) matches win first;
    otherwise a fuzzy containment match is attempted.
    """
    by_norm = {_normalize_header(c): c for c in available}
    mapping: Dict[str, Optional[str]] = {}
    used = set()

    for role, expected in required.items():
        target = _normalize_header(expected)
        if target in by_norm and by_norm[target] not in used:
            mapping[role] = by_norm[target]
            used.add(by_norm[target])
            continue

        # Fuzzy: expected contained in header or vice versa.
        found: Optional[str] = None
        for norm, original in by_norm.items():
            if original in used:
                continue
            if target in norm or norm in target:
                found = original
                break
        mapping[role] = found
        if found:
            used.add(found)
    return mapping


def validate_mapping(
    mapping: Dict[str, Optional[str]],
    required_roles: Dict[str, str],
    file_label: str,
) -> List[str]:
    """Return human-readable errors for missing required mappings."""
    errors: List[str] = []
    for role, expected in required_roles.items():
        chosen = mapping.get(role)
        if not chosen:
            errors.append(
                f"{file_label}: required column '{expected}' "
                f"(role '{role}') is not mapped."
            )
    return errors


def trim_text(value: Any) -> Optional[str]:
    """Trim whitespace; return None for null/blank."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, str) and value.strip() == "":
        return None
    text = str(value).strip()
    return text if text else None


def normalize_task_type(value: Any) -> Optional[str]:
    """Strip TaskType_ prefix; preserve meaningful category names."""
    text = trim_text(value)
    if text is None:
        return None
    prefix = TASK_TYPE_PREFIX
    if text.lower().startswith(prefix.lower()):
        text = text[len(prefix) :]
    return text.strip() or None


def normalize_complexity(value: Any) -> Optional[str]:
    """Strip Complexity_ prefix."""
    text = trim_text(value)
    if text is None:
        return None
    prefix = COMPLEXITY_PREFIX
    if text.lower().startswith(prefix.lower()):
        text = text[len(prefix) :]
    return text.strip() or None


def normalize_change_or_new(value: Any) -> Optional[str]:
    """Normalize New/Change to canonical 'New' or 'Change' (case-insensitive)."""
    text = trim_text(value)
    if text is None:
        return None
    key = text.lower()
    if key in CHANGE_OR_NEW_ALIASES:
        return CHANGE_OR_NEW_ALIASES[key]
    # Allow already-canonical values with mixed case.
    for alias, canonical in CHANGE_OR_NEW_ALIASES.items():
        if key == alias:
            return canonical
    return text  # keep unknown values as trimmed original for matching attempts


def normalize_work_category(value: Any) -> Optional[str]:
    """Normalize Work Category (trim only; no arbitrary stripping)."""
    return trim_text(value)


def normalize_difficulty(value: Any) -> Optional[str]:
    """Normalize Difficulty (trim only)."""
    return trim_text(value)


def parse_effort_manday(value: Any) -> Tuple[Optional[float], Optional[str]]:
    """
    Parse EffortManday into a numeric man-day value.

    Supports:
    - numeric values: 1, 1.5, 2.0
    - numeric strings: \"1\", \"1.5\"
    - labels: \"EffortManday_1\", \"EffortManday_1.5\"

    Returns (parsed_value, error_message). On success error_message is None.
    Missing/invalid values return (None, reason) and are never coerced to 0.
    """
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None, "Actual effort is missing."

    if isinstance(value, bool):
        return None, f"Actual effort value '{value}' is not a valid number."

    if isinstance(value, (int, float)):
        return float(value), None

    text = str(value).strip()
    if not text:
        return None, "Actual effort is blank."

    # Strip known categorical prefix if present.
    if text.lower().startswith(EFFORT_PREFIX.lower()):
        text = text[len(EFFORT_PREFIX) :]

    text = text.strip()
    if not text:
        return None, f"Actual effort label '{value}' has no numeric portion."

    try:
        return float(text), None
    except ValueError:
        return None, f"Actual effort '{value}' is not a valid number."


def parse_standard_effort(value: Any) -> Tuple[Optional[float], Optional[str]]:
    """
    Parse Standard Development Man-Day.

    Rejects missing, non-numeric, negative, and zero values.
    """
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None, "Standard effort is missing."

    if isinstance(value, bool):
        return None, f"Standard effort '{value}' is not a valid number."

    if isinstance(value, (int, float)):
        num = float(value)
    else:
        text = str(value).strip()
        if not text:
            return None, "Standard effort is blank."
        try:
            num = float(text)
        except ValueError:
            return None, f"Standard effort '{value}' is not a valid number."

    if num < 0:
        return None, f"Standard effort {num} is negative."
    if num == 0:
        return None, "Standard effort is zero and cannot be used as a reference."
    return num, None


_LEVEL_BOUNDARY = r"(?<![A-Za-z])"
_LEVEL_END = r"(?![A-Za-z])"


def parse_assignee(
    assignee: Any,
    coefficients: Optional[Dict[str, float]] = None,
) -> Dict[str, Any]:
    """
    Extract employee name and professional level from an Assignee string.

    Supported patterns (examples):
    - ``NGUYEN VAN A Associate/Department``
    - ``NGUYEN VAN B Specialist/Department``
    - ``LE DUC ANH Professional/(Part) Sales & CS``

    Levels are matched by explicit keyword (case-insensitive, whole-word).
    Unknown levels do not receive a default coefficient.
    """
    coeffs = coefficients or DEFAULT_COEFFICIENTS
    result: Dict[str, Any] = {
        "employee_name": None,
        "professional_level": None,
        "coefficient": None,
        "issues": [],
        "original_assignee": trim_text(assignee),
    }

    text = trim_text(assignee)
    if text is None:
        result["issues"].append("Assignee is blank.")
        result["issues"].append("BLANK_EMPLOYEE_NAME")
        return result

    # Prefer longer level names first to avoid partial overlaps.
    levels = sorted(coeffs.keys(), key=len, reverse=True)
    found_level: Optional[str] = None
    match_span: Optional[Tuple[int, int]] = None

    for level in levels:
        pattern = re.compile(
            _LEVEL_BOUNDARY + re.escape(level) + _LEVEL_END,
            re.IGNORECASE,
        )
        m = pattern.search(text)
        if m:
            found_level = level  # use canonical casing from config
            match_span = m.span()
            break

    if found_level is None:
        result["issues"].append("UNKNOWN_PROFESSIONAL_LEVEL")
        # Best-effort name: take text before first '/' if present.
        name_part = text.split("/")[0].strip()
        result["employee_name"] = name_part or None
        if not result["employee_name"]:
            result["issues"].append("BLANK_EMPLOYEE_NAME")
        return result

    result["professional_level"] = found_level
    result["coefficient"] = float(coeffs[found_level])

    # Name is everything before the matched level keyword.
    assert match_span is not None
    name = text[: match_span[0]].strip()
    # Clean trailing separators.
    name = re.sub(r"[\s\-/|]+$", "", name).strip()
    result["employee_name"] = name or None
    if not result["employee_name"]:
        result["issues"].append("BLANK_EMPLOYEE_NAME")
    return result


def build_lookup_key(
    category: Optional[str],
    change_or_new: Optional[str],
    difficulty: Optional[str],
) -> Optional[str]:
    """
    Build a case-insensitive lookup key.

    Uses upper-cased category/difficulty and canonical New/Change.
    """
    if not category or not change_or_new or not difficulty:
        return None
    return f"{category.strip().upper()}|{change_or_new.strip().upper()}|{difficulty.strip().upper()}"
