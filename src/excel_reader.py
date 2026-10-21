"""Excel / CSV input helpers."""

from __future__ import annotations

import io
from dataclasses import dataclass
from pathlib import Path
from typing import BinaryIO, List, Optional, Union

import pandas as pd
from openpyxl import load_workbook

PathLike = Union[str, Path, BinaryIO, bytes]


@dataclass
class WorkbookInfo:
    """Metadata about an uploaded workbook."""

    sheet_names: List[str]
    suggested_sheet: str
    is_csv: bool = False
    has_vba: bool = False
    filename: str = ""


class ExcelReadError(Exception):
    """Raised when a file cannot be read as a tabular workbook."""


def _to_bytes(source: PathLike) -> bytes:
    if isinstance(source, bytes):
        return source
    if hasattr(source, "read"):
        data = source.read()
        if hasattr(source, "seek"):
            try:
                source.seek(0)
            except Exception:
                pass
        if isinstance(data, str):
            return data.encode("utf-8")
        return data
    return Path(source).read_bytes()


def _filename_of(source: PathLike, fallback: str = "") -> str:
    if isinstance(source, (str, Path)):
        return Path(source).name
    name = getattr(source, "name", None)
    return name or fallback


def inspect_workbook(source: PathLike, filename: str = "") -> WorkbookInfo:
    """
    Inspect an uploaded file and return sheet names / suggestions.

    Does not execute macros. VBA presence is reported so the UI can warn users.
    """
    name = filename or _filename_of(source)
    lower = name.lower()

    if lower.endswith(".csv"):
        return WorkbookInfo(
            sheet_names=["CSV"],
            suggested_sheet="CSV",
            is_csv=True,
            has_vba=False,
            filename=name,
        )

    if not (lower.endswith(".xlsx") or lower.endswith(".xlsm") or lower.endswith(".xls")):
        raise ExcelReadError(
            f"Unsupported file type for '{name}'. "
            "Please upload .xlsx, .xlsm, or .csv."
        )

    data = _to_bytes(source)
    try:
        # keep_vba is not required for inspection; detect via extension / content.
        wb = load_workbook(io.BytesIO(data), read_only=True, data_only=False)
    except Exception as exc:  # noqa: BLE001 – surface as user-friendly error
        raise ExcelReadError(
            f"Could not open '{name}' as an Excel workbook. "
            "The file may be corrupted or password-protected."
        ) from exc

    sheets = list(wb.sheetnames)
    wb.close()
    if not sheets:
        raise ExcelReadError(f"Workbook '{name}' contains no worksheets.")

    suggested = _suggest_sheet(sheets, name)
    has_vba = lower.endswith(".xlsm")
    return WorkbookInfo(
        sheet_names=sheets,
        suggested_sheet=suggested,
        is_csv=False,
        has_vba=has_vba,
        filename=name,
    )


def _suggest_sheet(sheets: List[str], filename: str) -> str:
    """Pick a likely worksheet based on common naming conventions."""
    lowered = {s.lower(): s for s in sheets}
    hints = [
        "productivity",
        "report",
        "standard",
        "output",
        "detail",
        "sheet1",
        "data",
    ]
    fname = filename.lower()
    for hint in hints:
        if hint in fname:
            for key, original in lowered.items():
                if hint in key:
                    return original
    # Prefer Sheet1 when present.
    for key, original in lowered.items():
        if key == "sheet1":
            return original
    return sheets[0]


def read_table(
    source: PathLike,
    sheet_name: Optional[str] = None,
    filename: str = "",
) -> pd.DataFrame:
    """
    Read a worksheet (or CSV) into a DataFrame.

    - Preserves column order.
    - Does not execute Excel formulas; formula cells become NaN / cached values
      depending on openpyxl/pandas behaviour. Prefer data_only=False via
      pandas engine and coerce later.
    - Adds a 1-based ``_source_row`` column pointing at the Excel row number
      (header is row 1, first data row is 2).
    """
    name = filename or _filename_of(source)
    lower = name.lower()
    data = _to_bytes(source)

    try:
        if lower.endswith(".csv"):
            df = pd.read_csv(io.BytesIO(data))
        else:
            # engine=openpyxl; dtype=object keeps text labels intact.
            df = pd.read_excel(
                io.BytesIO(data),
                sheet_name=sheet_name or 0,
                engine="openpyxl",
                dtype=object,
            )
    except ExcelReadError:
        raise
    except Exception as exc:  # noqa: BLE001
        raise ExcelReadError(
            f"Failed to read data from '{name}'"
            + (f" (sheet '{sheet_name}')" if sheet_name else "")
            + f": {exc}"
        ) from exc

    if df is None or df.empty:
        raise ExcelReadError(
            f"Worksheet '{sheet_name or 'default'}' in '{name}' is empty."
        )

    # Normalize header names to strings and strip whitespace.
    df.columns = [str(c).strip() if c is not None else f"Unnamed_{i}" for i, c in enumerate(df.columns)]
    df = df.copy()
    df.insert(0, "_source_row", range(2, len(df) + 2))
    return df


def list_columns(df: pd.DataFrame) -> List[str]:
    """Return user-facing column names (excludes internal helper columns)."""
    return [c for c in df.columns if not str(c).startswith("_")]
