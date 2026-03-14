# Taiwan Label Print Macro

Automated per-SKU label printing for Taiwan warehouse operations in Excel.

![EXCEL](https://img.shields.io/badge/EXCEL-AUTOMATION-4d4d4d?style=for-the-badge)
![LANGUAGE](https://img.shields.io/badge/LANGUAGE-VBA-d4af37?style=for-the-badge)
![FEATURE](https://img.shields.io/badge/FEATURE-SKU_LABEL_PRINT-1e88e5?style=for-the-badge)
![WORKFLOW](https://img.shields.io/badge/WORKFLOW-WAREHOUSE_DISPATCH-43a047?style=for-the-badge)
![SELECTION](https://img.shields.io/badge/SELECTION-ROW_%2F_SKU_COLUMN-616161?style=for-the-badge)
![OUTPUT](https://img.shields.io/badge/OUTPUT-PRINT_PREVIEW-ff8f00?style=for-the-badge)
![STATUS](https://img.shields.io/badge/STATUS-STABLE-52c41a?style=for-the-badge)

A VBA macro that prints **one SKU label per page** from the warehouse dispatch sheet.

It is designed for a workflow where:

- warehouse staff prepare SKU-level dispatch data in Excel
- the user selects SKU rows or SKU cells
- the macro builds printable labels on a preview sheet
- each SKU is printed on its own label page
- the user confirms the layout in **Print Preview** before printing

---

## ✨ Features

### 1. One SKU = One Label

Each valid selected SKU generates **one separate label page**.

This is suitable for:

- product package labels
- Taiwan warehouse identification labels
- receiving and shelving support
- clerk-friendly SKU recognition

---

### 2. Selection-Aware Workflow

The macro supports several selection styles:

- selected rows
- selected SKU cells
- the whole SKU column

It automatically:

- ignores the header row
- ignores empty cells
- keeps only valid SKU rows

---

### 3. Print Preview First

The macro opens **Print Preview** before printing.

This allows the user to check:

- page layout
- label spacing
- font rendering
- SKU content
- missing / blank values

before sending labels to the printer.

---

### 4. Chinese / English Label Support

The printed label supports mixed Chinese and English fields such as:

- 货架位
- SKU
- GTIN
- 数量

Chinese label text is generated safely for Excel/VBA display.

---

### 5. Full GTIN Display

The macro forces GTIN to display as full text instead of scientific notation.

For example, it preserves:

```text
6942265706768
```

instead of printing:

```text
6.94227E+12
```

---

### 6. 100mm × 100mm Label Workflow

The macro is designed for square warehouse label paper such as:

```text
100mm × 100mm
```

with:

- one label per page
- clear border layout
- readable SKU text
- simple warehouse-facing format

---

## Label Content

Each label contains:

- **货架位**
- **SKU**
- **GTIN**
- **数量**

Example:

```text
货架位：A-01-04
SKU：MJ-魔仙-小野猫下睫毛
GTIN：
数量：110
```

If GTIN is empty, it remains blank.

---

## Expected Source Columns

The macro expects the source worksheet to contain at least these fields:

| Column | Field |
|---|---|
| A | SKU编号 |
| C | 台湾货架位 |
| F | 数量 |
| G | GTIN |

---

## How It Works

1. Open the workbook that contains the SKU dispatch data.
2. Go to the worksheet that contains **SKU编号 / 台湾货架位 / 数量 / GTIN**.
3. Select:
   - one or more rows, or
   - one or more SKU cells, or
   - the whole SKU column.
4. Run the macro.
5. The macro builds labels on `PrintPreview`.
6. It opens **Print Preview**.

---

## Design Notes

This macro was built for warehouse label printing where:

- one product SKU needs one label
- Taiwan shelf position must be visible
- some receiving clerks may not know the products well
- the label must help them identify and shelve correctly

It is optimized for:

- manual operational use
- SKU-level dispatch labeling
- warehouse receiving support
- repeatable label preview before print

---

## Recommended Use Case

Use this macro when you need to print **per-SKU labels** for:

- Taiwan warehouse transfer
- product package identification
- dispatch preparation
- receiving and shelving operations

---

## Status

**Stable**

The macro supports:

- one-label-per-page output
- selection-aware behavior
- Chinese / English labels
- GTIN full-number display
- print preview workflow

---

## File Purpose

This README documents the logic and usage of:

**Taiwan Label Print Macro**

for SKU-level warehouse label generation in Excel.
