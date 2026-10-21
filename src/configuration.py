"""Central configuration for Productivity Report Analyzer."""

from __future__ import annotations

from copy import deepcopy
from dataclasses import dataclass, field
from typing import Dict, List

# Professional level → coefficient (percent of specialist baseline).
DEFAULT_COEFFICIENTS: Dict[str, float] = {
    "Associate": 90.0,
    "Specialist": 100.0,
    "Professional": 143.0,
}

# Canonical New/Change values after normalization.
CHANGE_OR_NEW_ALIASES: Dict[str, str] = {
    "new": "New",
    "change": "Change",
}

# Known categorical prefixes that should be stripped during normalization.
TASK_TYPE_PREFIX = "TaskType_"
COMPLEXITY_PREFIX = "Complexity_"
EFFORT_PREFIX = "EffortManday_"

# Default column mappings: logical role → expected header name.
PRODUCTIVITY_REQUIRED_COLUMNS: Dict[str, str] = {
    "task_type": "TaskType",
    "change_or_new": "ChangeOrNew",
    "complexity": "Complexity",
    "effort_manday": "EffortManday",
    "assignee": "Assignee",
    "finished_date": "Finished Date",
}

STANDARD_REQUIRED_COLUMNS: Dict[str, str] = {
    "work_category": "Work Category",
    "new_change": "New/Change",
    "difficulty": "Difficulty",
    "standard_effort": "Standard Development Man-Day",
}

# Minimum output columns (Section 7).
OUTPUT_DETAIL_COLUMNS: List[str] = [
    "Employee Name",
    "Finished Date",
    "Complexity",
    "Total Effort (MD)",
    "Standard Effort",
    "Productivity",
    "Effort For Specialist",
]

# Extra audit / detail columns included in the default export.
OUTPUT_AUDIT_COLUMNS: List[str] = [
    "Task Key",
    "Task Type",
    "Change/New",
    "Professional Level",
    "Validation Issues",
]

# Display precision defaults.
DEFAULT_EFFORT_DECIMALS = 2
DEFAULT_PRODUCTIVITY_DECIMALS = 2

# Validation issue codes.
ISSUE_MISSING_STANDARD = "MISSING_STANDARD"
ISSUE_AMBIGUOUS_STANDARD = "AMBIGUOUS_STANDARD"
ISSUE_INVALID_ACTUAL_EFFORT = "INVALID_ACTUAL_EFFORT"
ISSUE_ZERO_ACTUAL_EFFORT = "ZERO_ACTUAL_EFFORT"
ISSUE_UNKNOWN_LEVEL = "UNKNOWN_PROFESSIONAL_LEVEL"
ISSUE_BLANK_EMPLOYEE = "BLANK_EMPLOYEE_NAME"
ISSUE_INVALID_STANDARD_EFFORT = "INVALID_STANDARD_EFFORT"
ISSUE_MISSING_CLASSIFICATION = "MISSING_CLASSIFICATION"


@dataclass
class AppConfig:
    """Runtime configuration that the UI may override."""

    coefficients: Dict[str, float] = field(
        default_factory=lambda: deepcopy(DEFAULT_COEFFICIENTS)
    )
    effort_decimals: int = DEFAULT_EFFORT_DECIMALS
    productivity_decimals: int = DEFAULT_PRODUCTIVITY_DECIMALS

    def level_names(self) -> List[str]:
        """Return professional level names sorted by coefficient descending."""
        return sorted(self.coefficients.keys(), key=lambda k: -self.coefficients[k])


def validate_coefficients(coefficients: Dict[str, float]) -> List[str]:
    """Return a list of validation error messages for the coefficient table."""
    errors: List[str] = []
    if not coefficients:
        errors.append("At least one professional level coefficient is required.")
        return errors
    for name, value in coefficients.items():
        if not str(name).strip():
            errors.append("Professional level name cannot be blank.")
        try:
            num = float(value)
        except (TypeError, ValueError):
            errors.append(f"Coefficient for '{name}' must be numeric.")
            continue
        if num < 0:
            errors.append(f"Coefficient for '{name}' cannot be negative.")
    return errors
