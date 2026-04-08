# Logistics PDF Renamer & RAR Packer

![Python](https://img.shields.io/badge/Python-3.9%2B-blue?logo=python&logoColor=white)
![Platform](https://img.shields.io/badge/Platform-Windows-0078D6?logo=windows&logoColor=white)
![Input](https://img.shields.io/badge/Input-PDF-red?logo=adobeacrobatreader&logoColor=white)
![Archive](https://img.shields.io/badge/Archive-WinRAR-orange)
![Encoding](https://img.shields.io/badge/Encoding-UTF--8-green)
![Mode](https://img.shields.io/badge/Mode-预刷-purple)
![Status](https://img.shields.io/badge/Status-Stable-brightgreen)

> **Detect Logistics → Count Labels → Rename PDFs → Archive to RAR → Clean Up**  
> Designed for stable, repeatable processing of Shopee logistics label PDFs on Windows.

---

## Overview

This script automates the processing of logistics label PDF files.

It will:

1. Detect which logistics type each PDF belongs to
2. Count how many labels are inside each PDF
3. Rename each PDF using a fixed naming rule
4. Pack the renamed PDFs into one dated `.rar` archive
5. Name the archive using the **total label count of the current run**
6. Save the archive into `Downloads`
7. Delete the processed PDF files after successful compression

The workflow is intended to be simple, repeatable, and suitable for daily use.

---

## Supported Logistics Types

Currently the script supports only these 3 types:

- **統一超商 / 交貨便** → `711`
- **標準配送 / 店到店** → `店到店`
- **FamilyMart / 全家** → `全家`

If a PDF does not belong to one of these three, the script will stop and show an error.

---

## Rename Rules

Each PDF is renamed to:

```text
目標名稱-單數.pdf
```

Examples:

```text
711-8单.pdf
店到店-152单.pdf
全家-5单.pdf
```

If there are multiple PDFs of the same logistics type, suffixes are added automatically:

```text
全家-5单(1).pdf
全家-3单(2).pdf
```

---

## Archive Naming Rule

After all PDFs are renamed, they are packed into one RAR file.

The archive name is based on:

- **today's date**
- fixed base name **`三得美（预刷）`**
- **the total label count of all PDFs in the current run**
- optional **batch suffix** when more than one batch is processed on the same day

### First batch of the day

```text
yyyymmdd三得美（预刷）-总单数单.rar
```

Example:

```text
20260408三得美（预刷）-179单.rar
```

### Second batch of the same day

```text
yyyymmdd三得美（预刷）-总单数单（第二批）.rar
```

Example:

```text
20260408三得美（预刷）-24单（第二批）.rar
```

### Third batch of the same day

```text
yyyymmdd三得美（预刷）-总单数单（第三批）.rar
```

Example:

```text
20260408三得美（预刷）-31单（第三批）.rar
```

### Batch detection rule

The script checks whether any same-day archive already exists with the same date and base family:

```text
yyyymmdd三得美（预刷）-
```

This check ignores the quantity of `单`.

So if `20260408三得美（预刷）-179单.rar` already exists, and a later run has only 24 labels, the new archive will still be:

```text
20260408三得美（预刷）-24单（第二批）.rar
```

---

## Output Location

The final RAR archive is saved to:

```text
C:\Users\zouzh\Downloads\
```

---

## Duplicate Order-ID Checking

The script checks duplicate `訂單編號` only when there are **multiple PDFs of the same logistics type**.

### Checked
- `711`
- `店到店`

### Skipped intentionally
- `FamilyMart / 全家`

Reason:

- FamilyMart PDFs often do not expose order IDs reliably through normal PDF text extraction
- this case is relatively rare
- manual checking is more practical than adding fragile OCR logic

When multiple FamilyMart PDFs are present, the script will still continue and print a warning that duplicate checking was skipped.

---

## Folder Structure

Place these files together in one folder:

```text
your_folder/
│
├─ logistics_pdf_renamer_and_rar.py
├─ run_logistics_pdf_renamer.bat
│
├─ 711_labels.pdf
├─ familymart_labels.pdf
└─ store_to_store_labels.pdf
```

You can name the PDFs anything you want.  
The script identifies them by **content**, not by filename.

---

## Processing Logic

For each `.pdf` in the script folder:

### 1. PDF Detection

The script reads text from the PDF and decides whether it is:

- `711`
- `店到店`
- `全家`

### 2. Label Counting

The script treats:

```text
1 page = 1 label
```

So the number of pages becomes the number of labels.

### 3. Duplicate Validation

If there are multiple PDFs of the same logistics type:

- `711` and `店到店` are checked for duplicated `訂單編號`
- `全家` is skipped intentionally

### 4. Rename

Each PDF is renamed according to logistics type and page count.

### 5. Calculate batch total

The script sums the page counts from **all PDFs in the current run**.

This total is used in the final RAR filename.

### 6. Archive

The script uses **WinRAR** to create a `.rar` archive.

### 7. Cleanup

Only after the RAR is created successfully, the processed PDFs are deleted.

---

## Safety & Error Handling

The script prints clear `ERROR:` messages and stops when something is wrong.

Protected cases include:

- no PDF files found
- unsupported logistics type
- ambiguous logistics detection
- failed PDF reading
- failed order-ID extraction for checked logistics types
- duplicated `訂單編號` in duplicated `711` / `店到店` PDFs
- target renamed file already exists
- WinRAR not found
- RAR creation failed
- failed deletion after archive creation

Example:

```text
ERROR: Unrecognized logistics type in file: abc123.pdf. Only these types are allowed: 統一超商 / 標準配送 / FamilyMart
ERROR: Duplicated 訂單編號 detected within duplicated logistics type PDFs.
```

---

## How to Run

### Use the `.bat` file

Double-click:

```text
run_logistics_pdf_renamer.bat
```

The runner will:

- switch to the script folder
- run the Python script
- keep the terminal open
- show whether the run succeeded or failed

---

## Dependency

Install Python package once:

```bash
pip install pypdf
```

You also need:

- **Python 3.9+**
- **WinRAR installed on Windows**

---

## Recommended Workflow

1. Put the PDF files into the same folder as the script
2. Make sure the PDFs are closed
3. Double-click the `.bat`
4. Check the generated RAR file in `Downloads`
5. Continue with the next batch

---

## Design Philosophy

- **Simple daily workflow**
- **No silent overwrites**
- **Stable handling of common cases**
- **Conservative error stopping**
- **Practical treatment of rare edge cases**

This script intentionally favors reliability over over-engineering.

---

## Troubleshooting Checklist

If something looks wrong:

- Is Python installed correctly?
- Is `pypdf` installed?
- Is WinRAR installed?
- Are the PDFs closed?
- Are the PDFs really one of the 3 supported logistics types?
- If there are multiple FamilyMart PDFs, remember that duplicate checking is skipped intentionally
- Is there already a same-day RAR batch in `Downloads`?

---

## End

If you want later enhancements such as log files, ZIP fallback, drag-and-drop launching, or summary reports, they can be added cleanly on top of this version.
