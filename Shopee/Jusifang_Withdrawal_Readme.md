# Jusifang Withdrawal Automation

![Python](https://img.shields.io/badge/Python-3.9%2B-blue?logo=python&logoColor=white)
![Platform](https://img.shields.io/badge/Platform-Windows-0078D6?logo=windows&logoColor=white)
![Excel](https://img.shields.io/badge/Input-Excel%20(.xlsx)-217346?logo=microsoft-excel&logoColor=white)
![OCR](https://img.shields.io/badge/OCR-Tesseract-success)
![Encoding](https://img.shields.io/badge/Encoding-UTF--8-green)
![Status](https://img.shields.io/badge/Status-Stable-brightgreen)

> **Template → OCR Process → Archive → Revert Template**  
> Designed for stable, repeatable withdrawal bookkeeping with screenshots embedded in Excel.

---

## Overview

This script automates the processing of **withdrawal screenshots embedded in Excel workbooks**.

It will:

1. Detect workbooks that contain **embedded screenshots** in column **`提款截圖`**
2. Use **OCR** to extract:
   - 日期
   - 提款編號
   - 提款金額
3. Fill missing cells **without overwriting existing data**
4. Generate a **dated, summed output workbook**
5. Archive the finished workbook into a **per-person folder**
6. Revert the working workbook back to its **original default template**

The workflow is safe, repeatable, and designed to minimize manual mistakes.

---

## Folder Structure (Final)

```
base_folder/
│
├─ Jusifang_Withdrawal_Automation.py
├─ Jusifang_Withdrawal_Automation.bat
│
├─ 個人卡提(馬京瑋).xlsx        ← working template (you paste images here)
├─ 個人卡提(白鎮瑋).xlsx        ← working template
│
├─ _template_sources/           ← auto-learned default templates (DO NOT edit)
│   ├─ 個人卡提(馬京瑋).xlsx
│   └─ 個人卡提(白鎮瑋).xlsx
│
├─ 個人卡提(馬京瑋)/            ← archived results
│   └─ 260122個人卡提(馬京瑋)5051.xlsx
│
├─ 個人卡提(白鎮瑋)/
│   └─ 260122個人卡提(白鎮瑋)120416.xlsx
│
└─ run_log/
    ├─ run_20260128_155915.txt
    └─ run_20260128_153459.txt
```

---

## Default Template Rules (Very Important)

A workbook is considered a **default template** if:

- It contains all required headers:
  - `日期`
  - `提款編號`
  - `提款金額`
  - `提款截圖`
- **No images** are embedded in column `提款截圖`

### What happens to templates

- On startup, the script **auto-learns** each clean template
- A copy is saved into `_template_sources/<核心檔名>.xlsx`
- After processing, the working workbook is **reverted** by copying from its own template source

This guarantees:

- `馬京瑋` always reverts to **馬京瑋’s default data**
- `白鎮瑋` always reverts to **白鎮瑋’s default data**
- No cross-contamination of default values

---

## Processing Logic

For each `.xlsx` in the base folder:

### 1. Template Detection
- If **no images** are found in `提款截圖` → skipped (treated as template)

### 2. OCR Processing
- Only rows with images in `提款截圖` are processed
- Existing cell values are **never overwritten**
- OCR extracts:
  - 日期 (YYYY/MM/DD)
  - 提款編號 (16–22 digits, OCR-robust)
  - 提款金額 (negative values → stored as positive)

### 3. Summary & Rename
- `newest_date` = latest 日期 in sheet
- `total_amount` = sum of 提款金額
- Output filename:
  ```
  yyMMdd + 核心檔名 + total_amount.xlsx
  ```

Example:
```
260122個人卡提(馬京瑋)5051.xlsx
```

### 4. Archive
- The finished workbook is moved to:
  ```
  ./核心檔名/
  ```

### 5. Revert Template
- The working file is deleted
- A fresh template is copied back from:
  ```
  _template_sources/<核心檔名>.xlsx
  ```

---

## Safety & Error Handling

The script prints clear `ERROR:` messages and **never overwrites silently**.

Common protected cases:

- Workbook is open in Excel
- Required headers are missing
- Finished workbook already exists
- Template source is missing
- Working file still contains images during revert

Example:
```
ERROR: Finished workbook already exists: 個人卡提(馬京瑋)/260122個人卡提(馬京瑋)5051.xlsx
ERROR: To avoid overwriting, the script will stop for this file.
```

---

## How to Run (Recommended)

### Use the `.bat` file (Windows)

Double-click:

```
Jusifang_Withdrawal_Automation.bat
```

The runner will:

- Force UTF-8 output
- Create `run_log/` automatically
- Capture **stdout + stderr** into a timestamped log
- Keep the terminal open
- Print the last 40 lines for quick checking

---

## Logs

- All logs are stored in:
  ```
  run_log/
  ```
- Encoding: **UTF-8**
- If terminal shows garbled characters:
  - The log file itself is correct
  - Open it with **VS Code** or **Windows 11 Notepad**
  - Terminal display depends on font / code page

---

## OCR Notes

- Uses **Tesseract OCR**
- Default language:
  ```
  chi_tra + eng
  ```
- Tesseract path (edit if needed):
  ```
  C:\Program Files\Tesseract-OCR\tesseract.exe
  ```

---

## Recommended Workflow (Daily Use)

1. Open `個人卡提(XXX).xlsx`
2. Paste withdrawal screenshots into column `提款截圖`
3. Save & close Excel
4. Double-click the `.bat`
5. Verify the archived file
6. Continue using the reverted clean template

---

## Design Philosophy

- **Template integrity first**
- **No silent overwrites**
- **Per-person isolation**
- **Logs over dialogs**
- **One-click repeatability**

This script is intentionally conservative to protect financial records.

---

## Troubleshooting Checklist

If something looks wrong:

- Is Excel fully closed?
- Does the workbook still contain images?
- Did a finished file already exist?
- Check `run_log/*.txt`

If needed, delete a wrong template source in `_template_sources/` and rerun once with a clean template to re-learn it.

---

## End

If you want enhancements later (versioned outputs, CSV export, summary report, auto-zip archives), they can be layered cleanly on top of this foundation.
