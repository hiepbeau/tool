"""Unit tests for column mapping and normalization."""

from __future__ import annotations

import pytest

from src.column_mapper import (
    auto_map_columns,
    build_lookup_key,
    normalize_change_or_new,
    normalize_complexity,
    normalize_task_type,
    parse_assignee,
    parse_effort_manday,
    parse_standard_effort,
    validate_mapping,
)
from src.configuration import (
    DEFAULT_COEFFICIENTS,
    PRODUCTIVITY_REQUIRED_COLUMNS,
    STANDARD_REQUIRED_COLUMNS,
)


def test_normalize_task_type_strips_prefix():
    assert normalize_task_type("TaskType_MI") == "MI"
    assert normalize_task_type("tasktype_ma") == "ma"
    assert normalize_task_type("  TaskType_WEB  ") == "WEB"
    assert normalize_task_type("MI") == "MI"  # no arbitrary stripping
    assert normalize_task_type(None) is None
    assert normalize_task_type("   ") is None


def test_normalize_complexity_strips_prefix():
    assert normalize_complexity("Complexity_C1") == "C1"
    assert normalize_complexity("complexity_c3") == "c3"
    assert normalize_complexity("C5") == "C5"


def test_normalize_change_or_new():
    assert normalize_change_or_new("Change") == "Change"
    assert normalize_change_or_new("change") == "Change"
    assert normalize_change_or_new("NEW") == "New"
    assert normalize_change_or_new(" new ") == "New"


def test_lookup_key_case_insensitive():
    k1 = build_lookup_key("MI", "Change", "C1")
    k2 = build_lookup_key("mi", "change", "c1")
    assert k1 == k2 == "MI|CHANGE|C1"


def test_auto_map_productivity_headers():
    available = [
        "Key",
        "Assignee",
        "TaskType",
        "ChangeOrNew",
        "Complexity",
        "EffortManday",
        "Finished Date",
    ]
    mapping = auto_map_columns(available, PRODUCTIVITY_REQUIRED_COLUMNS)
    assert mapping["task_type"] == "TaskType"
    assert mapping["assignee"] == "Assignee"
    assert mapping["finished_date"] == "Finished Date"
    errors = validate_mapping(mapping, PRODUCTIVITY_REQUIRED_COLUMNS, "Productivity Report")
    assert errors == []


def test_auto_map_standard_headers():
    available = [
        "Work Category",
        "Work Type",
        "New/Change",
        "Difficulty",
        "Standard Development Man-Day",
    ]
    mapping = auto_map_columns(available, STANDARD_REQUIRED_COLUMNS)
    assert mapping["work_category"] == "Work Category"
    assert mapping["standard_effort"] == "Standard Development Man-Day"
    assert validate_mapping(mapping, STANDARD_REQUIRED_COLUMNS, "Standard") == []


def test_missing_column_reported():
    mapping = auto_map_columns(["Assignee"], PRODUCTIVITY_REQUIRED_COLUMNS)
    errors = validate_mapping(mapping, PRODUCTIVITY_REQUIRED_COLUMNS, "Productivity Report")
    assert any("TaskType" in e for e in errors)


def test_standard_effort_rejects_zero_and_negative():
    assert parse_standard_effort(0)[0] is None
    assert parse_standard_effort(-1)[0] is None
    value, err = parse_standard_effort(2.5)
    assert value == pytest.approx(2.5)
    assert err is None


def test_effort_missing_not_zero():
    value, err = parse_effort_manday(None)
    assert value is None
    assert err is not None


def test_assignee_word_boundary_avoids_partial_match():
    # "Associate" must not match inside "Associateship" (trailing letters).
    info = parse_assignee("NGUYEN Associateship/Dept", DEFAULT_COEFFICIENTS)
    assert info["professional_level"] is None
    assert "UNKNOWN_PROFESSIONAL_LEVEL" in info["issues"]
