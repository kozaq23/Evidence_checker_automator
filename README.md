# Evidence_checker_automator
Evidence checker automation checks the Excel exercises __Music Data__ and __Excel Functions__



## 🔍 File Identification

Function: `identify_wp3_file(path)`

- Loads an Excel file and inspects the structure.
- Classification:
  - `"music"`: ≥400 rows and columns `Year`, `Album`, `Artist`, `Total Sales`.
  - `"dashboard"`: ≥90 rows and includes `Name`, `Date`, `Department`, `Rating`.
  - `"error"`: If file cannot be opened (e.g. still open in Excel).

---

## 📑 Sheet Selection

Function: `select_best_sheet(excel)`

- Prioritises sheets not named `RAW DATA` or `TASK ONE` unless only one sheet exists.
- Detection logic:
  - **Dashboard**: sheet with `Rating` column and >90 non-null values.
  - **Music**: last column is numeric and has ≥400 values.
- Fallback to `RAW DATA` or `TASK ONE` if no matches.

---

## ✅ Shared Checks (All File Types)

### `check_nulls(df)`
- Detects presence of any blank cells.
- Returns affected row numbers.

### `check_duplicates(df)`
- Identifies fully duplicated rows.
- Returns indices of duplicates.

---

## 🎵 Music Dataset Checks

### `check_artist_column(df)`
- Checks that:
  - No leading/trailing whitespace in `Artist`.
  - Capitalisation uses `.title()`.
- Reports each row with issues.

### `check_album_duplicates(df)`
- Ensures `"Greatest Hits"` appears **more than once** in the `Album` column.

### `check_total_sales(df, path, sheet)`
- Verifies:
  - Sum of `Total Sales` equals expected figure.
  - Cells formatted in:
    - **GBP** currency (`£`)
    - **Two decimal places**

### `check_table_format(df, path, sheet)`
- Confirms use of Excel *Table* (structured data):
  - Must start at **A1**
  - Match range of the dataset exactly
  - Only one table per sheet allowed

---

## 📊 Dashboard Dataset Checks

### `check_qs(df, path, sheet)`
- Validates `"QS"` replaced with `"Quality Surveyor"`.
- Accepts if ≥16 instances found.
- Detects regex-based misspellings and casing issues.

### `check_validation(df, path, sheet)`
- Checks Excel **data validation** is applied:
  - `"Rating"` column → `"whole number"`
  - `"Department"` column → `"list"`

### `check_functions(df, path, sheet)`
- Ensures presence of Excel formulas:
  - Required: `SUM`, `MAX`, `MIN`, `AVERAGE`, `MEDIAN`, `MODE`, `STDEV.S`
  - Accepts alternative: `STDEV` instead of `STDEV.S`
- Parses cell formula strings.

---

## 📝 Logging and Tracking

### `log_startup(config)`
- Records:
  - Timestamp
  - Python, pandas, openpyxl versions
  - Hostname
  - UI config state

### `log_shutdown(start_ts, run_id, path, messages, exit_code)`
- Logs:
  - File analysed
  - Summary of all message types (info, ok, error)
  - Failed checks
  - Function invocation count (`check_counts`)
  - Execution time

---

## 🖥️ User Interface Features

- Built with `tkinter`
- Supports:
  - Dark mode toggle
  - File selection (auto/manual)
  - Live path display
  - Status bar feedback
  - Message filtering: show/hide info, ok, error
- Tooltips embedded for all major controls

---
