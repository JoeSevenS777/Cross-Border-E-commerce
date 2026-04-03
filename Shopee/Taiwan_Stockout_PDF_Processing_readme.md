# Instock Logistics PDF Renamer & RAR Packer

![Python](https://img.shields.io/badge/Python-3.9%2B-blue?logo=python&logoColor=white)
![Platform](https://img.shields.io/badge/Platform-Windows-0078D6?logo=windows&logoColor=white)
![Input](https://img.shields.io/badge/Input-PDF-red?logo=adobeacrobatreader&logoColor=white)
![Archive](https://img.shields.io/badge/Archive-WinRAR-orange)
![Mode](https://img.shields.io/badge/Mode-现货-purple)
![Status](https://img.shields.io/badge/Status-Stable-brightgreen)

> **Detect Logistics → Count Only Logistics Labels → Create No-Pick Copy → Archive to RAR → Clean Up**  
> Built for mixed logistics PDFs where one logistics label may be followed by one or more pick-label pages.

---

## Overview

This script processes **instock / 现货** logistics PDFs.

It will:

1. Detect the logistics type of each PDF
2. Detect which pages are **logistics labels**
3. Ignore **pick-label pages** and **empty pages**
4. Rename the original PDF by the count of logistics labels
5. Create a second PDF that keeps **only logistics-label pages**
6. Rename that second PDF with the suffix `"(无拣货单)"`
7. Pack all generated PDFs into one dated `.rar` archive
8. Save the archive into `Downloads`
9. Delete the processed PDFs after successful compression

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
- or 3+ pick-label pages
- and sometimes even empty pages

So page count is not reliable.

Instead, the script classifies each page as:

- **logistics label**
- **pick label**
- **empty page**

Then it counts only the **logistics-label pages**.

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

After processing, all generated PDFs are packed into one RAR archive named:

```text
yyyymmdd三得美（现货）.rar
```

Example:

```text
20260403三得美（现货）.rar
```

If the same-day archive already exists in `Downloads`, the script creates:

```text
20260403三得美（现货）(2).rar
20260403三得美（现货）(3).rar
```

and so on.

---

## Output Location

The final RAR archive is saved to:

```text
C:\Users\zouzh\Downloads\
```

---

## Page Classification Logic

The script classifies every page using text patterns.

### Pick-label pages
Usually contain markers like:

- `Order No:`
- `Tracking No:`
- `Total Items:`
- `货架位`
- `数量`

### Empty pages
Pages with no meaningful text are treated as empty.

### Logistics-label pages
Pages matching carrier-specific shipping-label features are treated as logistics pages.

### Fallback
If a page is **not clearly a pick label** and **not empty**, the script treats it as a logistics label.

This follows the practical rule:

> Detect non-logistics pages first, then subtract them.

---

## Folder Structure

Place these files together in one folder:

```text
your_folder/
│
├─ Logistics Pdf Renamer Instock.py
├─ Logistics Pdf Renamer Instock.bat
│
├─ 711-6個.pdf
├─ 蝦皮店到店-134個.pdf
├─ 全家-1個.pdf
├─ 隔日達-12個.pdf
└─ 萊爾富-2個.pdf
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
- empty page

### 3. Count only logistics labels
Only logistics-label pages are counted.

### 4. Rename original PDF
The original full PDF is renamed based on logistics type and logistics-label count.

### 5. Create no-pick copy
A second PDF is generated containing only logistics-label pages.

### 6. Archive
All processed PDFs are packed into a `.rar` archive using WinRAR.

### 7. Cleanup
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
Logistics Pdf Renamer Instock.bat
```

The runner will:

- switch to the script folder
- run the Python script
- keep the terminal open
- show whether the run succeeded or failed

If your Python filename contains spaces, the `.bat` should call it with quotes.

Example:

```bat
python "Logistics Pdf Renamer Instock.py"
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
- **Detect non-logistics pages first**
- **Keep only real logistics labels in the no-pick copy**
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

---

## End

This version is designed for daily instock PDF processing where one logistics label may be followed by multiple pick labels or empty pages.
