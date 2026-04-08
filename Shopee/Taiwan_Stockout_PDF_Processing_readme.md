# Instock Logistics PDF Renamer & RAR Packer

![Python](https://img.shields.io/badge/Python-3.9%2B-blue?logo=python&logoColor=white)
![Platform](https://img.shields.io/badge/Platform-Windows-0078D6?logo=windows&logoColor=white)
![Input](https://img.shields.io/badge/Input-PDF-red?logo=adobeacrobatreader&logoColor=white)
![Archive](https://img.shields.io/badge/Archive-WinRAR-orange)
![Mode](https://img.shields.io/badge/Mode-现货-purple)
![Status](https://img.shields.io/badge/Status-Stable-brightgreen)

> **Detect Logistics → Count Only Real Logistics Labels → Create No-Pick Copy → Archive to RAR → Clean Up**  
> Built for mixed logistics PDFs where one logistics label may be followed by one or more pick-label pages, including multi-page pick lists.

---

## Overview

This script processes **instock / 现货** logistics PDFs.

It will:

1. Detect the logistics type of each PDF
2. Detect which pages are **real logistics labels**
3. Ignore **pick-label pages**, **pick-label continuation pages**, and **empty pages**
4. Rename the original PDF by the count of real logistics labels
5. Create a second PDF that keeps **only logistics-label pages**
6. Rename that second PDF with the suffix `"(无拣货单)"`
7. Pack all generated PDFs into one dated `.rar` archive
8. Name the archive by the **total logistics-label count of the whole batch**
9. Save the archive into `Downloads`
10. Delete the processed PDFs after successful compression

---

## Supported Logistics Types

Currently the script supports these 5 types:

- **統一超商 / 交貨便** → `711`
- **標準配送 / 店到店** → `店到店`
- **FamilyMart / 全家** → `全家`
- **隔日到貨** → `隔日到货`
- **Hi-Life** → `莱尔福`

If a PDF does not belong to one of these five, the script stops and shows an error.

---

## Why This Script Is Different

This script does **not** count labels by page count.

That is because a PDF may contain:

- 1 logistics label page
- followed by 1 pick-label page
- or 2 pick-label pages
- or a pick-label table that spills into another page
- and sometimes even empty pages

So page count is not reliable.

Instead, the script uses **sequence-aware page classification** and counts only the **real logistics-label pages**.

It can now distinguish between:

- **logistics label**
- **pick label**
- **pick continuation**
- **empty page**
- **unknown page**

This is specifically designed to avoid over-counting when a pick list is split across multiple pages.

---

## Rename Rules

### Original PDF

Each original PDF is renamed to:

```text
目标名称-单数.pdf
```

Examples:

```text
711-6单.pdf
店到店-134单.pdf
隔日到货-12单.pdf
莱尔福-2单.pdf
全家-1单.pdf
```

### No-pick-label PDF

A second PDF is created that removes:

- all pick-label pages
- all pick-label continuation pages
- all empty pages

It is renamed to:

```text
目标名称-单数(无拣货单).pdf
```

Examples:

```text
711-6单(无拣货单).pdf
店到店-134单(无拣货单).pdf
隔日到货-12单(无拣货单).pdf
莱尔福-2单(无拣货单).pdf
全家-1单(无拣货单).pdf
```

### Duplicate logistics-type PDFs

If multiple PDFs map to the same target name, suffixes are added automatically:

```text
店到店-20单(1).pdf
店到店-15单(2).pdf
店到店-20单(1)(无拣货单).pdf
店到店-15单(2)(无拣货单).pdf
```

---

## Archive Naming Rule

After processing, all generated PDFs are packed into one RAR archive.

The archive name is based on:

- **today's date**
- fixed base name **`三得美（现货）`**
- **the total count of all real logistics labels in the current run**
- optional **batch suffix** when more than one batch is processed on the same day

### First batch of the day

```text
yyyymmdd三得美（现货）-总单数单.rar
```

Example:

```text
20260408三得美（现货）-4单.rar
```

### Second batch of the same day

```text
yyyymmdd三得美（现货）-总单数单（第二批）.rar
```

Example:

```text
20260408三得美（现货）-2单（第二批）.rar
```

### Third batch of the same day

```text
yyyymmdd三得美（现货）-总单数单（第三批）.rar
```

Example:

```text
20260408三得美（现货）-4单（第三批）.rar
```

### Batch detection rule

The script checks whether any same-day archive already exists with the same date and base family:

```text
yyyymmdd三得美（现货）-
```

This check ignores the quantity of `单`.

So if `20260408三得美（现货）-4单.rar` already exists, and a later run has only 2 orders, the new archive will still be:

```text
20260408三得美（现货）-2单（第二批）.rar
```

---

## Output Location

The final RAR archive is saved to:

```text
C:\Users\zouzh\Downloads\
```

---

## Page Classification Logic

The script classifies every page using text patterns **and page order context**.

### 1. Logistics-label pages

Pages matching carrier-specific shipping-label features are treated as logistics pages.

These are the **anchor pages** used for counting real orders.

### 2. Pick-label pages

Usually contain markers like:

- `Order No:`
- `Tracking No:`
- `Total Items:`
- `货架位`
- `數量`
- `数量`

### 3. Pick-continuation pages

Some pick lists continue onto the next page and may no longer contain the usual header markers.

These continuation pages are detected by features such as:

- repeated shelf-location row patterns like `E-05-01`
- item-table row structure
- previous page being a pick page or another continuation page
- absence of strong logistics-label markers

This prevents a second page of pick items from being mistaken as a logistics label.

### 4. Empty pages

Pages with no meaningful text are treated as empty.

### 5. Unknown pages

If a page is not clearly logistics, pick, continuation, or empty, the script classifies it conservatively as **unknown**, not logistics.

This prevents accidental over-counting.

---

## Folder Structure

Place these files together in one folder:

```text
your_folder/
│
├─ logistics_pdf_renamer_instock.py
├─ logistics_pdf_renamer_instock.bat
│
├─ any_name_1.pdf
├─ any_name_2.pdf
├─ any_name_3.pdf
└─ ...
```

The PDF filenames themselves do not matter.  
The script identifies them by **content**, not by filename.

---

## Processing Logic

### 1. Detect logistics type

The script scans the PDF text and decides whether the file is:

- `711`
- `店到店`
- `全家`
- `隔日到货`
- `莱尔福`

### 2. Classify pages

Each page is classified as:

- logistics label
- pick label
- pick continuation
- empty page
- unknown page

### 3. Count only logistics labels

Only real logistics-label pages are counted.

### 4. Rename original PDF

The original full PDF is renamed based on logistics type and logistics-label count.

### 5. Create no-pick copy

A second PDF is generated containing only logistics-label pages.

### 6. Calculate batch total

The script sums the logistics-label counts from **all PDFs in the current run**.

This total is used in the final RAR filename.

### 7. Archive

All processed PDFs are packed into a `.rar` archive using WinRAR.

### 8. Cleanup

After successful archive creation, the processed PDFs in the working folder are deleted.

---

## Safety & Error Handling

The script stops with clear `ERROR:` messages when something goes wrong.

Protected cases include:

- no PDF files found
- unsupported logistics type
- ambiguous logistics detection
- failed PDF reading
- no logistics pages detected
- failed page extraction
- failed rename
- failed no-pick PDF creation
- WinRAR not found
- RAR creation failed
- failed deletion after archive creation

---

## How to Run

Double-click:

```text
logistics_pdf_renamer_instock.bat
```

The runner will:

- switch to the script folder
- run the Python script
- keep the terminal open
- show whether the run succeeded or failed

If your Python filename contains spaces, the `.bat` should call it with quotes.

Example:

```bat
python "logistics_pdf_renamer_instock.py"
```

---

## Dependencies

Install once:

```bash
pip install pypdf
```

You also need:

- **Python 3.9+**
- **WinRAR installed on Windows**

---

## Recommended Workflow

1. Put the PDFs into the same folder as the script
2. Make sure the PDFs are closed
3. Double-click the `.bat`
4. Check the generated archive in `Downloads`
5. Continue with the next batch

---

## Design Philosophy

- **Do not trust page count**
- **Detect real logistics anchors**
- **Recognize multi-page pick lists**
- **Keep only real logistics labels in the no-pick copy**
- **Prevent uncertain pages from inflating the order count**
- **Prefer practical stability over rigid assumptions**
- **Stop clearly when something is truly wrong**

---

## Troubleshooting Checklist

If something looks wrong:

- Is Python installed correctly?
- Is `pypdf` installed?
- Is WinRAR installed?
- Are the PDFs closed?
- Is the `.bat` pointing to the exact `.py` filename?
- Are the files really one of the 5 supported logistics types?
- Did the PDF contain readable text, not just images?
- Is a pick list continuing onto a second page?
- Is there already a same-day RAR batch in `Downloads`?

---

## End

This version is designed for daily instock PDF processing where one logistics label may be followed by multiple pick labels, multi-page pick-label continuations, or empty pages.
