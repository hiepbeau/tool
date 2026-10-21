# Productivity Report Analyzer

Compare an employee **Productivity Report** against a customer **Standard** effort table, calculate task-level productivity, review validation issues, and export an Excel report.

Designed for non-technical users on Windows: upload Excel files, confirm column mappings, calculate, and download results.

## Features

- Upload Productivity Report, Standard (kept in session — upload once), and optional Output Template
- Worksheet selection for multi-sheet workbooks
- Editable column mapping with required-column validation
- Normalization of `TaskType_`, `Complexity_`, and New/Change values
- Robust EffortManday parsing (`1`, `"1.5"`, `EffortManday_1`, …)
- Professional level detection (Associate / Specialist / Professional) with configurable coefficients
- Task-level productivity and Effort For Specialist
- Validation issues for missing/ambiguous standards, invalid effort, unknown levels
- Overall **weighted** productivity (Σ standard ÷ Σ actual × 100)
- Excel export with detail, summary, validation, and standard reference sheets

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
   - **Productivity Report** — employee tasks (required each run)
   - **Standard** — approved standard effort table (upload once; kept until replaced)
   - **Output Template** — optional layout for the export
2. **Select worksheets** if a workbook has more than one sheet.
3. **Confirm column mapping** — defaults match the sample files; override if your headers differ.
4. **Review data preview** — row counts, sample rows, detected levels.
5. **Adjust coefficients / decimal places** if needed (defaults: Associate 90, Specialist 100, Professional 143).
6. Click **Calculate Productivity**.
7. Review the detail table and validation issues.
8. Click **Download Excel Report**.

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
