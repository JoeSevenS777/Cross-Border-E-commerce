# 🔗 Shopee Settlement Screenshot Migration Automation (Python)

![Excel](https://img.shields.io/badge/Excel-Workbook-555555)
![Python](https://img.shields.io/badge/Python-3.10%2B-3776AB)
![Platform](https://img.shields.io/badge/Platform-Shopee-EE4D2D)
![OCR](https://img.shields.io/badge/OCR-Tesseract-6A5ACD)
![Status](https://img.shields.io/badge/Status-Production-0078D4)

This Python automation reads embedded Shopee settlement screenshot images from the `pics` sheet, extracts settlement dates, account identity, and payout amounts, then migrates the results into the `reconciliation` sheet.

It automatically appends only new settlement batches, preserves formulas and formats, copies payment-related columns `J:N`, leaves confirmation column `O` blank for manual checking, deletes processed images, saves a timestamped workbook, and moves the previous workbook into a `log` folder.

---

## Features

- Reads embedded screenshot images from the `pics` sheet.
- Identifies wallet/account by bank information in the screenshot.
- Supports current account mapping:

  | Bank information in screenshot | Reconciliation account |
  |---|---|
  | `華南商業銀行 (7494)` | `ck1808052` |
  | `中華郵政股份有限公司 (6260)` | `noso1981` |

- Extracts:
  - Start date
  - End date
  - Revenue / payout amount
- Sorts image results by date order before writing.
- Appends new rows to the `reconciliation` sheet.
- Updates existing rows if the same `account + start date + end date` already exists.
- Copies formulas and formatting like Excel drag-down.
- Merges columns `J:O` by the same batch period.
- Copies/fills `J:N` from the previous batch.
- Leaves merged column `O` empty for manual confirmation.
- Marks start/end date cells red if the date period is not continuous.
- Deletes processed images from the `pics` sheet after successful migration.
- Creates a new timestamped workbook, for example:

      Shop1&4reconciliation_20260509163225.xlsx

- Moves the previous workbook into the `log` folder.
- Automatically detects the newest workbook matching:

      Shop1&4reconciliation*.xlsx

---

## Folder Structure

Put the following files in the same folder:

    京京withdrawal - 副本
    ├── migrate_sheet_pics_to_reconciliation.py
    ├── Run_Migrate_Sheet_Pics_To_Reconciliation_UPDATED.bat
    ├── Shop1&4reconciliation.xlsx
    └── log

After running the script, the folder may become:

    京京withdrawal - 副本
    ├── migrate_sheet_pics_to_reconciliation.py
    ├── Run_Migrate_Sheet_Pics_To_Reconciliation_UPDATED.bat
    ├── Shop1&4reconciliation_20260509163225.xlsx
    └── log
        └── Shop1&4reconciliation.xlsx

The newest timestamped workbook becomes the working workbook for the next run.

---

## Workbook Requirements

The workbook should contain these two sheets:

| Sheet name | Purpose |
|---|---|
| `pics` | Stores the embedded screenshots to be processed |
| `reconciliation` | Stores the settlement reconciliation table |

The script can also recognize common aliases such as:

    Pics
    Sheet Pics
    Reconciliation
    Sheet Reconciliation

However, using exactly these names is recommended:

    pics
    reconciliation

---

## Reconciliation Sheet Layout

The script expects these columns:

| Column | Header / Meaning | Script behavior |
|---|---|---|
| A | `account` | Filled directly |
| B | `start date` | Filled directly |
| C | `end date` | Filled directly |
| D | `revenue` | Filled directly |
| E:I | Formula columns | Copied down from previous row |
| J:N | Transfer/payment-related columns | Copied from previous batch |
| O | `confirm` | Merged but left blank |

The script treats rows with the same `start date + end date` as one batch.

For example:

| account | start date | end date | revenue |
|---|---:|---:|---:|
| ck1808052 | 2026/1/27 | 2026/2/2 | 87297 |
| noso1981 | 2026/1/27 | 2026/2/2 | 57268 |

These two rows belong to the same batch, so columns `J:O` are merged vertically across those two rows.

---

## Date Continuity Rule

The new batch start date should follow the previous batch end date.

Example:

    Previous end date: 2026/1/26
    New start date   : 2026/1/27

This is valid.

If the new start date does not equal:

    previous end date + 1 day

then the script fills the `start date` and `end date` cells in red.

This helps you quickly notice missing periods or incorrect screenshots.

---

## How the Script Reads Screenshots

Each screenshot should show records like:

    2026-02-02
    自動提款
    至帳號：華南商業銀行 (7494)
    支出
    -41,250

The script extracts all dates and negative amounts from one image.

For each image:

    start date = earliest date in the image
    end date   = latest date in the image
    revenue    = absolute sum of all negative amounts

Example:

    -41,250
    -10,813
    -13,805
    -10,362
    -11,067

The revenue becomes:

    41,250 + 10,813 + 13,805 + 10,362 + 11,067 = 87,297

---

## Required Software

### 1. Python

Install Python 3.10 or later.

During installation, tick:

    Add Python to PATH

Check Python:

    python --version

### 2. Required Python Packages

The `.bat` file checks these packages automatically:

    openpyxl
    pillow
    pytesseract

If missing, it tries to install them automatically.

Manual installation command:

    python -m pip install openpyxl pillow pytesseract

### 3. Tesseract OCR Engine

The script uses Tesseract OCR through `pytesseract`.

You must install the Tesseract OCR engine separately.

Recommended language support:

    English
    Chinese Traditional

If the Chinese Traditional OCR package is unavailable, the script falls back to English OCR.

---

## How to Use

### Step 1 — Put files in the same folder

Make sure these files are together:

    migrate_sheet_pics_to_reconciliation.py
    Run_Migrate_Sheet_Pics_To_Reconciliation_UPDATED.bat
    Shop1&4reconciliation.xlsx

### Step 2 — Put screenshots into the `pics` sheet

Insert or paste the screenshots into the `pics` sheet.

Each image should correspond to one account and one settlement period.

### Step 3 — Double-click the `.bat` file

Run:

    Run_Migrate_Sheet_Pics_To_Reconciliation_UPDATED.bat

The command window will show:

- Selected workbook
- Workbook sheet names
- Detected account
- Detected start date
- Detected end date
- Detected revenue
- Detected amounts

### Step 4 — Check the new workbook

The script creates a new timestamped workbook:

    Shop1&4reconciliation_yyyymmddhhmmss.xlsx

The previous workbook is moved to:

    log

Use the newest timestamped workbook next time.

---

## Output Example

The command window may show:

    Detected image results:
    ================================================================================
    Image 1
      Account   : ck1808052
      Start date: 2026/01/27
      End date  : 2026/02/02
      Revenue   : 87,297
      Dates     : 2026/01/27, 2026/01/28, 2026/01/29, 2026/01/30, 2026/02/02
      Amounts   : 41,250, 10,813, 13,805, 10,362, 11,067
    --------------------------------------------------------------------------------
    Image 2
      Account   : noso1981
      Start date: 2026/01/27
      End date  : 2026/02/02
      Revenue   : 57,268
      Dates     : 2026/01/27, 2026/01/28, 2026/01/29, 2026/01/30, 2026/02/02
      Amounts   : 27,526, 9,557, 7,102, 7,246, 5,837
    --------------------------------------------------------------------------------

---

## Important Notes

- Do not open the workbook in Excel while running the script.
- The script deletes images from the `pics` sheet only after successful migration.
- The old workbook is not deleted; it is moved to the `log` folder.
- Column `O` is intentionally left blank because confirmation should be done manually.
- If a new bank account appears, update the `ACCOUNT_MAP` section in the Python script.

Example:

    ACCOUNT_MAP = {
        "華南商業銀行 (7494)": "ck1808052",
        "中華郵政股份有限公司 (6260)": "noso1981",
        "New Bank Text (1234)": "new_account_name",
    }

---

## Troubleshooting

### Error: Cannot find workbook

Make sure the workbook name starts with:

    Shop1&4reconciliation

and ends with:

    .xlsx

Example valid names:

    Shop1&4reconciliation.xlsx
    Shop1&4reconciliation_20260509163225.xlsx

### Error: Cannot find required sheet

Check that the workbook has these sheets:

    pics
    reconciliation

### Error: Cannot identify account

The OCR result did not match the bank account mapping.

Open the Python script and add the new bank text to:

    ACCOUNT_MAP

### Error: Cannot extract dates or amounts

Possible causes:

- Screenshot is too blurry.
- Screenshot is cropped.
- Amounts are not shown as negative numbers.
- OCR did not recognize the text correctly.

Use clearer screenshots and try again.

---

## Current Script Files

    migrate_sheet_pics_to_reconciliation.py
    Run_Migrate_Sheet_Pics_To_Reconciliation_UPDATED.bat

---

## Version

    Version: 1.0
    Status : Production
