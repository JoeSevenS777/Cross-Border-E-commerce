# Taiwan Outer Package Label Macro - Active Workbook Version

Automated outer-package label printing for exported packing workbooks in Excel.

![EXCEL](https://img.shields.io/badge/EXCEL-AUTOMATION-4d4d4d?style=for-the-badge)
![LANGUAGE](https://img.shields.io/badge/LANGUAGE-VBA-d4af37?style=for-the-badge)
![FEATURE](https://img.shields.io/badge/FEATURE-OUTER_PACKAGE_LABEL-1e88e5?style=for-the-badge)
![WORKFLOW](https://img.shields.io/badge/WORKFLOW-ACTIVE_WORKBOOK-43a047?style=for-the-badge)
![GROUPING](https://img.shields.io/badge/GROUPING-BOX_NUMBER-616161?style=for-the-badge)
![OUTPUT](https://img.shields.io/badge/OUTPUT-PRINT_PREVIEW-ff8f00?style=for-the-badge)
![STATUS](https://img.shields.io/badge/STATUS-STABLE-52c41a?style=for-the-badge)

A VBA macro that prints **outer-carton labels** based on selected **箱号** cells in an exported picking & packing workbook.

It is designed for a workflow where:

- the **master workbook** stores the macro
- the **packing workbook** is opened separately for the picking clerk
- the user selects one or more cells in the **箱号** column
- the macro gathers **all rows for each selected box number**
- one outer-package label is generated per box, with pagination when needed

---

## ✨ Features

### 1. Active Workbook Workflow

The macro is stored in the **master workbook**, but works on the **currently active workbook and active worksheet**.

This matches the real workflow:

- export the packing workbook
- open that workbook
- select box numbers there
- run the macro
- preview the labels

---

### 2. Box-Based Grouping

The macro does **not** print row by row.

Instead, it groups all matching rows by **箱号**.

For example, if you select one cell containing:

`01`

the macro will still gather **all rows with 箱号 01** across the worksheet.

---

### 3. Header and Blank Exclusion

The macro ignores:

- the header row
- empty cells in the 箱号 column
- rows without valid SKU / quantity information

---

### 4. Preserved Box Number Format

Box numbers are normalized and displayed as:

- `01`
- `02`
- `03`

instead of:

- `1`
- `2`
- `3`

---

### 5. Multi-Page Support for Large Boxes

If one box contains too many SKU lines to fit on a single 100mm × 100mm label page, the macro automatically continues onto additional pages.

Each page repeats the header information:

- 箱号
- SKU数
- 明细SKU * 数量

So continuation pages remain readable and complete.

---

### 6. Print Preview First

The macro opens **Print Preview** before printing.

This allows the user to check layout, pagination, and content before sending labels to the printer.

---

## Label Content

Each outer-package label contains:

- **箱号**
- **SKU数**
- **明细SKU * 数量**

Example:

```text
箱号：01
SKU数：2
明细SKU * 数量：
MJ-萌睫尚品-MJS03  *  30
MJ-萌睫尚品-大脸猫大容量  *  10
```

---

## Expected Source Columns

The macro expects the active worksheet to contain at least these columns:

| Column | Field |
|---|---|
| A | SKU编号 |
| H | 实发数量 |
| I | 箱号 |

---

## How It Works

1. Open the exported **picking & packing workbook**.
2. Go to the worksheet that contains **SKU编号 / 实发数量 / 箱号**.
3. Select one or more cells in the **箱号** column.
4. Run the macro.
5. The macro finds all matching rows for those box numbers.
6. It generates a paginated outer-package label layout on `PrintPreview`.
7. It opens **Print Preview**.

---

## Design Notes

This macro was built for warehouse dispatch workflows where:

- goods are packed into cartons by **箱号**
- one carton can contain multiple SKUs
- the outer label must summarize carton contents clearly
- the master workbook and clerk workbook are separated

It is optimized for:

- exported operational workbooks
- manual selection by box number
- clear carton identification during dispatch and receiving

---

## Recommended Use Case

Use this macro when you need to print **outer-carton content labels** after actual packing quantities and box numbers have been written back into the exported packing workbook.

---

## Status

**Stable**

The macro supports:

- active workbook execution
- grouped carton labels
- repeated headers on continuation pages
- print preview workflow
- box-number normalization

---

## File Purpose

This README documents the logic and usage of:

**Taiwan Outer Package Label Macro - Active Workbook Version**

for warehouse dispatch and carton-label generation in Excel.
