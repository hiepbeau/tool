# Productivity Report Analyzer

Compare one or more employee **Productivity Reports** against a shared customer **Standard** effort table, calculate task-level productivity, review validation issues, and export each report as its own Excel workbook.

Designed for non-technical users on Windows: upload Excel files, confirm column mappings, calculate, browse each report, and download results.

## Features

- Upload **multiple** Productivity Reports in one batch, plus one shared Standard (kept in session) and optional Output Template
- Independent productivity calculation per source file, plus combined totals across all files
- Browse individual reports with a modern slider / prev–next interface
- Filter and review all task-level results (including Source File)
- Worksheet selection for multi-sheet workbooks
- Editable column mapping with required-column validation across all uploaded reports
- Normalization of `TaskType_`, `Complexity_`, and New/Change values
- Robust EffortManday parsing (`1`, `"1.5"`, `EffortManday_1`, …)
- Professional level detection (Associate / Specialist / Professional) with configurable coefficients
- Task-level productivity and Effort For Specialist
- Validation issues for missing/ambiguous standards, invalid effort, unknown levels — tagged by source file
- Excel export: download **each report individually**

## Requirements

- Python 3.11+
- Windows (also works on macOS/Linux)

## Setup (Windows)

```powershell
cd "C:\Users\LG CNS\Desktop\AI_PJ\CompareTool"
python -m venv .venv
.\.venv\Scripts\activate
pip install -r requirements.txt
```

## Run the application

```powershell
streamlit run app.py
```

Then open the URL shown in the terminal (usually http://localhost:8501).

## Run tests

```powershell
pytest -q
```

## How to use

1. **Upload files**
   - **Productivity Report(s)** — one or many employee task files
   - **Standard** — shared approved standard effort table (upload once; kept until replaced)
   - **Output Template** — optional layout for the export
2. **Select worksheets** for each uploaded report (and Standard / Template).
3. **Confirm column mapping** — applies to all reports; each file must contain the mapped headers.
4. **Review data preview** — per-file row counts and sample rows.
5. **Adjust coefficients / decimal places** if needed ().
6. Click **Calculate Productivity** — runs independently per file and builds a combined summary.
7. **Browse by report** with the slider, or open **All results combined** to filter every task.
8. Download each report individually from the Export section.

## Calculation formulas

| Metric | Formula |
|---|---|
| Productivity (%) | Standard Effort (MD) ÷ Actual Effort (MD) × 100 |
| Effort For Specialist (MD) | Employee Coefficient × Actual Effort (MD) ÷ 100 |
| Overall weighted productivity | Σ Standard Effort ÷ Σ Actual Effort × 100 (valid tasks only) |

Example: MI + Change + C1, actual 1 MD, standard 2 MD, Professional (143) → **200%** productivity, **1.43 MD** Effort For Specialist.

## Sample files

Real samples are in `samples/`:

| File | Role |
|---|---|
| `Productivity Report.xlsx` | Task list (`TaskType`, `ChangeOrNew`, `Complexity`, `EffortManday`, `Assignee`, `Finished Date`, …) |
| `Standard.xlsx` | Lookup table (`Work Category`, `New/Change`, `Difficulty`, `Standard Development Man-Day`) |
| `Output_Template.xlsx` | Expected export columns including Employee Name, Complexity, efforts, Productivity, Effort For Specialist |

## Project structure

```
CompareTool/
├── app.py                 # Streamlit UI
├── requirements.txt
├── README.md
├── samples/
├── src/
│   ├── configuration.py
│   ├── excel_reader.py
│   ├── column_mapper.py
│   ├── validators.py
│   ├── productivity_calculator.py
│   └── report_generator.py
└── tests/
```

## CSV note

CSV uploads are supported for Productivity Report and Standard. CSV does **not** preserve Excel formatting or multiple worksheets.

## Limitations

- `.xlsm` macros are never executed; VBA is not preserved on export.
- Output Template columns that have no source mapping (e.g.) are left blank; calculated fields are written into matching headers.
- The primary detail sheet is **one row per source task** (not aggregated), even if the sample Output Template shows aggregated Task Count rows.
- Productivity is exported as a **percentage** (e.g. 200), per the business rules — not as a 0–1 ratio.
- Password-protected workbooks are not supported.
