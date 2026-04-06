# Taiwan Outer Package Label Macro - Active Workbook Version (Updated)

Automated outer-package label printing for exported packing workbooks in Excel.

![EXCEL](https://img.shields.io/badge/EXCEL-AUTOMATION-4d4d4d?style=for-the-badge)
![LANGUAGE](https://img.shields.io/badge/LANGUAGE-VBA-d4af37?style=for-the-badge)
![FEATURE](https://img.shields.io/badge/FEATURE-OUTER_PACKAGE_LABEL-1e88e5?style=for-the-badge)
![WORKFLOW](https://img.shields.io/badge/WORKFLOW-ACTIVE_WORKBOOK-43a047?style=for-the-badge)
![GROUPING](https://img.shields.io/badge/GROUPING-BOX_NUMBER-616161?style=for-the-badge)
![OUTPUT](https://img.shields.io/badge/OUTPUT-PRINT_PREVIEW-ff8f00?style=for-the-badge)
![STATUS](https://img.shields.io/badge/STATUS-WORKING-52c41a?style=for-the-badge)

---

## 🔄 Current Version Summary

This version is the **active-workbook print-preview edition** of the Taiwan outer-package label macro.

It now supports:

- **dynamic header detection**
- **selection-based processing**
- **whole-column smart selection**
- **strict row validation**
- **box-based grouping**
- **per-box page numbering**
- **14 detail rows per page**
- **warehouse shelf shown in a dedicated right column**
- **PrintPreview output before printing**

---

## ✅ Confirmed Layout Behavior

The generated label now follows the target layout in the workbook sample sheet:

### Row 1
- Left: `美妝補貨`
- Right: page text such as `第1頁/共5頁`

### Row 2
- Left: `SKU總數： 62`
- Right: `箱號: 02`

### Row 3
- Left: `明細SKU * 數量↓`
- Right: `貨架位↓`

### Rows 4+
- Column A: `SKU * 數量`
- Column B: `貨架位`

### Formatting Notes
- Font family: **Microsoft JhengHei**
- Page text font size: **8**
- `箱號:` stays on **one line**
- WrapText is disabled for label cells
- Column B is used as a dedicated right-side shelf column

---

## 🔍 Header Detection Rules

The macro does **not** rely on fixed column positions.

| Field | Accepted Header |
| --- | --- |
| SKU | `SKU编号` |
| Quantity | `实发数量` |
| Box No | `箱號` / `箱号` |
| Shelf | `台湾货架位` |

### Important
For the current version, shelf location is detected **only** by:

- `台湾货架位`

It does **not** search for `货架位`.

---

## 🧠 Selection Logic

When selecting cells in the **箱号 column**:

### Case A — Normal selection
- Only the selected non-empty 箱号 rows are processed

### Case B — Whole column selection
- Treated as **all non-empty 箱号 rows within the data range**
- Empty rows are ignored automatically

This matches the real workflow:

> Select the whole 箱号 column = process all existing boxes

---

## ✅ Validation Rules

Validation is applied only to the rows actually being processed.

| Field | Rule | Result |
| --- | --- | --- |
| SKU编号 | cannot be empty | stop |
| 实发数量 | cannot be empty | stop |
| 实发数量 | must be numeric | stop |
| 实发数量 | must be > 0 | stop |
| 实发数量 | must be integer | stop |
| 台湾货架位 | cannot be empty | stop |

### Example warning
```text
台湾货架位 为空（行 25，箱号 03）。请修正后再打印。
```

### Validation behavior
- blank rows are ignored
- only real selected rows are validated
- the first invalid row stops the process immediately

---

## 📦 Box-Based Output Logic

### 1. Grouping
Rows are grouped by **箱号**.

### 2. One label series per box
Each box produces its own label pages.

### 3. Box number normalization
Numeric box numbers are normalized like:

- `1` → `01`
- `2` → `02`
- `3` → `03`

### 4. Per-box page numbering
Page numbering is calculated **inside each box**, not across all boxes.

#### Example
If box `01` has only 1 page:
- `第1頁/共1頁`

If box `02` has 5 pages:
- `第1頁/共5頁`
- `第2頁/共5頁`
- `第3頁/共5頁`
- `第4頁/共5頁`
- `第5頁/共5頁`

---

## 📄 Pagination Rules

### Detail rows per page
The current version uses:

- **14 detail rows per page**

### Important pagination behavior
- page numbering resets for each box
- total pages are calculated per box
- there is **no blank spacer row between pages**
- this avoids the earlier issue where every content page was followed by an empty print page

---

## 🧾 Label Example

```text
美妝補貨                          第1頁/共2頁
SKU總數： 2                      箱號: 01

明細SKU * 數量↓                  貨架位↓
MJ-AAA * 30                      F-05-01
MJ-BBB * 10                      F-05-02
```

---

## 🧩 Workflow

1. Open the exported packing workbook
2. Activate the data sheet
3. Select cells in the **箱号** column  
   - or select the whole 箱号 column
4. Run the macro
5. Validation runs on the applicable rows
6. Labels are generated into `PrintPreview`
7. Excel opens **Print Preview**

---

## ⚠️ Error Handling Summary

| Scenario | Behavior |
| --- | --- |
| required header missing | stop and show missing headers |
| selection is not a range | stop |
| no valid non-empty 箱号 found | stop |
| invalid selected row | stop and show row number + box number |
| whole column selected | ignore blanks, validate only real rows |

---

## 📌 Current Status

**Working (Per-Box Pagination Layout Version)**

This version has been tested and adjusted to solve:

- layout matching to the sample sheet
- title/header placement
- per-box page numbering
- 14-row page capacity
- empty-page issue caused by spacer rows between page blocks

---

## 🧠 Practical Assessment

For current warehouse use, this version is:

- **usable**
- **stable enough for daily operation**
- **not yet deeply refactored**

That means it is suitable for real work, but still has room for code cleanup and structural improvement.

---

## 🚀 Recommended Next Upgrade

Recommended future improvements:

- refactor duplicated validation logic into one routine
- replace repeated box searching with a dictionary-based structure
- centralize all layout constants
- add a full error handler for unexpected runtime failures
- optionally show all invalid rows at once instead of stopping at the first one

---

## 📎 Reference

Updated from the earlier README provided by the user.

---

**End of README**
