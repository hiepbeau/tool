"""Unit tests for productivity calculation rules."""

from __future__ import annotations

import pandas as pd
import pytest

from src.column_mapper import parse_assignee, parse_effort_manday
from src.configuration import AppConfig, DEFAULT_COEFFICIENTS
from src.productivity_calculator import calculate_productivity
from src.validators import build_standard_lookup


def _prod_row(**overrides):
    base = {
        "_source_row": 2,
        "Key": "T-1",
        "TaskType": "TaskType_MI",
        "ChangeOrNew": "Change",
        "Complexity": "Complexity_C1",
        "EffortManday": "EffortManday_1",
        "Assignee": "LE DUC ANH Professional/(Part) Sales & CS",
        "Finished Date": "2026-10-15",
    }
    base.update(overrides)
    return base


def _std_row(**overrides):
    base = {
        "_source_row": 2,
        "Work Category": "MI",
        "New/Change": "Change",
        "Difficulty": "C1",
        "Standard Development Man-Day": 2,
    }
    base.update(overrides)
    return base


PROD_MAP = {
    "task_type": "TaskType",
    "change_or_new": "ChangeOrNew",
    "complexity": "Complexity",
    "effort_manday": "EffortManday",
    "assignee": "Assignee",
    "finished_date": "Finished Date",
}

STD_MAP = {
    "work_category": "Work Category",
    "new_change": "New/Change",
    "difficulty": "Difficulty",
    "standard_effort": "Standard Development Man-Day",
}


def _run(prod_rows, std_rows, config=None):
    prod_df = pd.DataFrame(prod_rows)
    std_df = pd.DataFrame(std_rows)
    lookup = build_standard_lookup(std_df, STD_MAP)
    return calculate_productivity(prod_df, PROD_MAP, lookup, config or AppConfig()), lookup


# ----- Test 1: Standard calculation -----


def test_standard_calculation_example():
    result, _ = _run([_prod_row()], [_std_row()])
    row = result.detail.iloc[0]
    assert row["Employee Name"] == "LE DUC ANH"
    assert row["Complexity"] == "C1"
    assert row["Total Effort (MD)"] == pytest.approx(1.0)
    assert row["Standard Effort"] == pytest.approx(2.0)
    assert row["Productivity"] == pytest.approx(200.0)
    assert row["Effort For Specialist"] == pytest.approx(1.43)
    assert row["Professional Level"] == "Professional"
    assert row["Employee Coefficient"] == pytest.approx(143.0)


# ----- Test 2: Different professional levels -----


@pytest.mark.parametrize(
    "level,expected",
    [
        ("Associate", 1.8),
        ("Specialist", 2.0),
        ("Professional", 2.86),
    ],
)
def test_effort_for_specialist_by_level(level, expected):
    assignee = f"NGUYEN VAN A {level}/Department"
    result, _ = _run(
        [_prod_row(Assignee=assignee, EffortManday="EffortManday_2")],
        [_std_row()],
    )
    row = result.detail.iloc[0]
    assert row["Effort For Specialist"] == pytest.approx(expected)
    assert row["Professional Level"] == level


# ----- Test 3: Missing standard -----


def test_missing_standard():
    result, _ = _run(
        [_prod_row(TaskType="TaskType_ZZ")],
        [_std_row()],
    )
    row = result.detail.iloc[0]
    assert len(result.detail) == 1  # task preserved
    assert pd.isna(row["Standard Effort"]) or row["Standard Effort"] is None
    assert pd.isna(row["Productivity"]) or row["Productivity"] is None
    assert "MISSING_STANDARD" in row["Validation Issues"]


# ----- Test 4: Zero actual effort -----


def test_zero_actual_effort_no_division():
    result, _ = _run(
        [_prod_row(EffortManday="EffortManday_0")],
        [_std_row()],
    )
    row = result.detail.iloc[0]
    assert row["Total Effort (MD)"] == pytest.approx(0.0)
    assert pd.isna(row["Productivity"]) or row["Productivity"] is None
    assert "ZERO_ACTUAL_EFFORT" in row["Validation Issues"]


# ----- Test 5: Invalid actual effort -----


def test_invalid_actual_effort_not_zero():
    value, err = parse_effort_manday("EffortManday_X")
    assert value is None
    assert err is not None

    result, _ = _run(
        [_prod_row(EffortManday="EffortManday_X")],
        [_std_row()],
    )
    row = result.detail.iloc[0]
    assert row["Total Effort (MD)"] is None or pd.isna(row["Total Effort (MD)"])
    assert "INVALID_ACTUAL_EFFORT" in row["Validation Issues"]
    assert pd.isna(row["Productivity"]) or row["Productivity"] is None


# ----- Test 6: Duplicate / conflicting standard key -----


def test_ambiguous_standard_key():
    result, lookup = _run(
        [_prod_row()],
        [
            _std_row(**{"Standard Development Man-Day": 2, "_source_row": 2}),
            _std_row(**{"Standard Development Man-Day": 5, "_source_row": 3}),
        ],
    )
    assert lookup.ambiguous_keys
    row = result.detail.iloc[0]
    assert "AMBIGUOUS_STANDARD" in row["Validation Issues"]
    assert pd.isna(row["Standard Effort"]) or row["Standard Effort"] is None
    assert pd.isna(row["Productivity"]) or row["Productivity"] is None


def test_identical_duplicate_standard_consolidated():
    result, lookup = _run(
        [_prod_row()],
        [
            _std_row(**{"Standard Development Man-Day": 2, "_source_row": 2}),
            _std_row(**{"Standard Development Man-Day": 2, "_source_row": 3}),
        ],
    )
    assert "MI|CHANGE|C1" in lookup.duplicate_identical_keys
    row = result.detail.iloc[0]
    assert row["Standard Effort"] == pytest.approx(2.0)
    assert row["Productivity"] == pytest.approx(200.0)


# ----- Test 7: Case and whitespace normalization -----


def test_case_and_whitespace_normalization():
    result, _ = _run(
        [
            _prod_row(
                TaskType="  tasktype_mi  ",
                ChangeOrNew=" change ",
                Complexity=" complexity_c1 ",
            )
        ],
        [
            _std_row(
                **{
                    "Work Category": " mi ",
                    "New/Change": "CHANGE",
                    "Difficulty": " c1 ",
                }
            )
        ],
    )
    row = result.detail.iloc[0]
    assert row["Task Type"] == "mi" or row["Task Type"].upper() == "MI"
    assert row["Standard Effort"] == pytest.approx(2.0)
    assert row["Productivity"] == pytest.approx(200.0)


# ----- Test 8: Unknown professional level -----


def test_unknown_professional_level():
    result, _ = _run(
        [_prod_row(Assignee="LE DUC ANH Intern/Sales")],
        [_std_row()],
    )
    row = result.detail.iloc[0]
    assert row["Productivity"] == pytest.approx(200.0)  # still calculated
    assert row["Effort For Specialist"] is None or pd.isna(row["Effort For Specialist"])
    assert "UNKNOWN_PROFESSIONAL_LEVEL" in row["Validation Issues"]
    assert row["Employee Coefficient"] is None or pd.isna(row["Employee Coefficient"])


# ----- Effort parser extras -----


@pytest.mark.parametrize(
    "raw,expected",
    [
        (1, 1.0),
        (1.5, 1.5),
        ("1", 1.0),
        ("1.5", 1.5),
        ("EffortManday_1", 1.0),
        ("EffortManday_1.5", 1.5),
    ],
)
def test_effort_parser_variants(raw, expected):
    value, err = parse_effort_manday(raw)
    assert err is None
    assert value == pytest.approx(expected)


def test_assignee_parsing_patterns():
    for text, name, level in [
        ("NGUYEN VAN A Associate/Department", "NGUYEN VAN A", "Associate"),
        ("NGUYEN VAN B Specialist/Department", "NGUYEN VAN B", "Specialist"),
        ("LE DUC ANH Professional/(Part) Sales & CS", "LE DUC ANH", "Professional"),
    ]:
        info = parse_assignee(text, DEFAULT_COEFFICIENTS)
        assert info["employee_name"] == name
        assert info["professional_level"] == level
        assert info["coefficient"] == DEFAULT_COEFFICIENTS[level]


def test_weighted_productivity_not_simple_average():
    """Σ standard / Σ actual, not average of percentages."""
    result, _ = _run(
        [
            _prod_row(EffortManday="EffortManday_1", Key="A"),  # prod 200%
            _prod_row(
                EffortManday="EffortManday_2",
                Complexity="Complexity_C2",
                Key="B",
                _source_row=3,
            ),  # std for MI Change C2 = need add
        ],
        [
            _std_row(),
            _std_row(
                Difficulty="C2",
                **{"Standard Development Man-Day": 3, "_source_row": 3},
            ),
        ],
    )
    # Task1: 2/1=200%, Task2: 3/2=150%. Average would be 175.
    # Weighted: (2+3)/(1+2)*100 = 166.666...
    assert result.summary.weighted_productivity == pytest.approx(500 / 3)
