# Picklist Shelf Replacer for Logistics PDF

![Python](https://img.shields.io/badge/Python-3.9%2B-blue)
![Platform](https://img.shields.io/badge/Platform-Windows-1f6feb)
![Input](https://img.shields.io/badge/Input-PDF-orange)
![Data](https://img.shields.io/badge/Data-XLSM-2ea44f)
![Mode](https://img.shields.io/badge/Mode-Pick%20List-purple)
![Output](https://img.shields.io/badge/Output-Updated%20PDF-red)
![Status](https://img.shields.io/badge/Status-Stable-brightgreen)

> **Match SKU → fetch 台湾货架位 from workbook → replace the Shelf value in pick-list pages only**  
> Built for mixed logistics PDFs that contain both **logistics labels** and **pick-list pages**.

---

## Overview

This script processes logistics-label PDFs and updates the **Shelf** column inside the **pick list**.

It will:

1. Detect **pick-list pages** by locating the table headers **Shelf / SKU / QTY / GTIN**
2. Skip normal logistics-label pages
3. Read the workbook **台湾货架位明细表_含打印预览（现版本）.xlsm**
4. Match the PDF **SKU** against workbook column **SKU编号**
5. Fetch the matching value from workbook column **台湾货架位**
6. Replace the original **Shelf** value with the new shelf code
7. Preserve the rest of the page as much as possible
8. Output a new updated PDF in the same folder
9. Print the full processing log in the terminal

---

## Workbook Mapping Rule

The script uses this relationship:

```text
PDF pick-list SKU  →  Excel SKU编号  →  Excel 台湾货架位
```

Example:

```text
JZ-黑夹  →  M06-05-01
MJ-BQI-10排宽梗鱼尾免胶假睫毛;单盒  →  M02-01-03
```

---

## Current Behavior

- **Only pick-list pages are modified**
- **Only the Shelf text is replaced**
- Shelf codes are written **on one line (no wrapping)**
- If a SKU cannot be matched:
  - the row is treated as unmatched
  - the script reports it in the terminal
  - depending on the current script version, unmatched row content may be blanked visually

---

## Folder Setup

Put these files into the same folder:

```text
D:\JoeProgramFiles\Replace Picklist Shelf in Logistics Label
```

Example:

```text
D:\JoeProgramFiles\Replace Picklist Shelf in Logistics Label
│
├─ Shelf Replacing.py
├─ Run Replace Picklist Shelf.bat
└─ your_logistics_pdf.pdf
```

The workbook stays here:

```text
D:\JoeProgramFiles\BigSeller库存管理\台湾货架位明细表_含打印预览（现版本）.xlsm
```

---

## Configuration

Inside the Python script, the main settings are:

```python
WORKBOOK_INPUT = r"D:\JoeProgramFiles\BigSeller库存管理\台湾货架位明细表_含打印预览（现版本）.xlsm"
WORKBOOK_SKU_COLUMN = "SKU编号"
WORKBOOK_SHELF_COLUMN = "台湾货架位"
```

If your workbook columns ever change, update those values.

---

## How It Works

### 1. Detect the input PDF
The script automatically picks the newest `.pdf` file in the same folder as the script.

### 2. Read the workbook
It opens the `.xlsm` file and builds a lookup table:

```text
SKU编号 → 台湾货架位
```

### 3. Detect pick-list pages
A page is treated as a pick list if it contains:

```text
Shelf
SKU
QTY
GTIN
```

### 4. Extract each pick-list row
For every row in the table, the script reads:

- Shelf
- SKU
- QTY

### 5. Match SKU to the workbook
Matching priority:

1. exact match after normalization
2. PDF SKU contained in workbook SKU
3. workbook SKU contained in PDF SKU

### 6. Replace the Shelf value
If matched, the script writes the new shelf code over the original shelf area.

### 7. Save output PDF
The script creates a file like:

```text
output_台湾货架位_originalfilename.pdf
```

---

## Run the Script

### Option 1: Run by BAT file
Double-click:

```text
Run Replace Picklist Shelf.bat
```

This will:

- check Python
- check required packages
- find the Python script automatically
- run the processing script
- keep the terminal open so you can see the log

### Option 2: Run by command line

```bash
python "Shelf Replacing.py"
```

You can also pass custom paths:

```bash
python "Shelf Replacing.py" "input.pdf" "workbook.xlsm" "output.pdf"
```

---

## Terminal Output Example

```text
Workbook sheet: Sheet2
Workbook SKU column: SKU编号
Workbook shelf column: 台湾货架位
Loaded SKU → 台湾货架位 records: 228
Page 2: REPLACE | SKU=JZ-黑夹 | A-07 -> M06-05-01 | exact
Page 2: REPLACE | SKU=MJ-BQI-10排宽梗鱼尾免胶假睫毛;单盒 | Z-01 -> M02-01-03 | exact
Page 4: REPLACE | SKU=MJ-萌睫尚品-狐系下睫毛02 | I-05 -> M05-03-10 | exact
```

---

## Requirements

Install Python packages:

```bash
pip install pymupdf openpyxl
```

Used libraries:

- **PyMuPDF** (`fitz`) — read and edit PDF text areas
- **openpyxl** — read the Excel workbook

---

## Notes

- This script is designed for **Windows**
- It is intended for your current logistics-label / pick-list PDF format
- Different PDF layouts may require adjustment of text detection logic
- If the replacement still looks slightly misaligned, font size or cover width can be fine-tuned in the script

---

## Troubleshooting

### The BAT says it cannot find the Python script
Make sure the `.bat` file and the `.py` file are in the same folder.

### The workbook SKU column cannot be found
Check that this is correct:

```python
WORKBOOK_SKU_COLUMN = "SKU编号"
```

### The replacement PDF looks messy
Use the latest script version that replaces **only the shelf text** and keeps **no wrap**.

### A SKU was not replaced
Possible reasons:

- the SKU does not exist in the workbook
- the SKU text in the PDF differs from the workbook
- the match is ambiguous

Check the terminal log for details.

---

## Recommended Script Version

Use the latest **no-wrap** version of the script for the cleanest result:

```text
Shelf Replacing.py
```

with the behavior:

- replace only shelf words
- keep code on one line
- preserve the original table as much as possible

---

## Summary

This tool is for quickly converting old pick-list shelf values like:

```text
A-07
Z-01
I-05
```

into real Taiwan shelf locations like:

```text
M06-05-01
M02-01-03
M05-03-10
```

based on your master workbook.
