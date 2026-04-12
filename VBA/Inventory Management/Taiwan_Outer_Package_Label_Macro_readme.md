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

## Current Version Summary

This version is the active-workbook print-preview edition of the Taiwan outer-package label macro.

It currently supports:

- dynamic header detection
- selection-based processing
- whole-column smart selection
- strict row validation
- box-based grouping
- per-box page numbering
- per-box SKU count display
- per-box total quantity display
- 14 detail rows per page
- dedicated columns for shelf, SKU, and quantity
- PrintPreview output before printing

---

## Confirmed Layout Behavior

The generated label now follows this 3-column layout:

### Row 1
- `A:B` merged: `美妝補貨`
- `C`: `箱號: 01`

### Row 2
- `A:B` merged: `SKU總數： 2  |  總數量：65`
- `C`: `第1頁/共1頁`

### Row 3
- `A`: `貨架位↓`
- `B`: `明細SKU ↓`
- `C`: `數量↓`

### Rows 4+
- Column A: shelf location
- Column B: SKU
- Column C: quantity

### Current formatting notes
- font family: **Microsoft JhengHei**
- title font size: **16**
- box number font size: **12**
- page info font size: **8**
- header row font size: **13**
- detail row font size: **12**
- WrapText is disabled
- row 2 now combines **SKU總數** and **總數量** in the merged `A:B` area
- page numbering is centered in column C row 2
- box number is centered in column C row 1

---

## Header Detection Rules

The macro does not rely on fixed column positions.

| Field | Accepted Header |
| --- | --- |
| SKU | `SKU编号` |
| Quantity | `实发数量` |
| Box No | `箱號` / `箱号` |
| Shelf | `台湾货架位` |

### Important
For the current version, shelf location is detected only by:

- `台湾货架位`

It does not search for `货架位`.

---

## Selection Logic

When selecting cells in the box-number column:

### Case A — Normal selection
- Only the selected non-empty 箱号 rows are processed

### Case B — Whole column selection
- Treated as all non-empty 箱号 rows within the data range
- Empty rows are ignored automatically

This matches the real workflow:

> Select the whole 箱号 column = process all existing boxes

---

## Validation Rules

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

## Quantity Summary Behavior

### Per-box total quantity
The macro now calculates **total quantity by box**.

That total is displayed together with the SKU count in row 2:

- `SKU總數： x  |  總數量：y`

### Example
If one box contains quantities:

- `10`
- `10`
- `10`
- `10`
- `10`
- `10`
- `5`

Then row 2 will display:

- `SKU總數： 7  |  總數量：65`

---

## Box-Based Output Logic

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
Page numbering is calculated inside each box, not across all boxes.

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

## Pagination Rules

### Detail rows per page
The current version uses:

- **14 detail rows per page**

### Important pagination behavior
- page numbering resets for each box
- total pages are calculated per box
- there is no blank spacer row between pages
- the earlier blank-page issue has been corrected by removing the page gap

---

## Label Example

```text
美妝補貨                          箱號: 01
SKU總數： 2  |  總數量：60        第1頁/共1頁

貨架位↓      明細SKU ↓            數量↓
D-04-04      MJ-萌睫尚品-MJ3D      10
F-05-04      JZ-夹子收纳盒         50
```

---

## Workflow

1. Open the exported packing workbook
2. Activate the data sheet
3. Select cells in the **箱号** column
   - or select the whole 箱号 column
4. Run the macro
5. Validation runs on the applicable rows
6. Labels are generated into `PrintPreview`
7. Excel opens Print Preview

---

## Current Layout Constants

The current VBA version uses these layout settings:

| Setting | Value |
| --- | --- |
| rows per page | `14` |
| page gap | `0` |
| column A width | `11.5` |
| column B width | `35.5` |
| column C width | `7.6` |
| title / SKU-total row height | `24` |
| header row height | `22.1` |
| detail row height | `20.05` |

---

## Error Handling Summary

| Scenario | Behavior |
| --- | --- |
| required header missing | stop and show missing headers |
| selection is not a range | stop |
| no valid non-empty 箱号 found | stop |
| invalid selected row | stop and show row number + box number |
| whole column selected | ignore blanks, validate only real rows |

---

## Current Status

**Working (3-Column Shelf-SKU-Qty Layout Version)**

This version has been adjusted to support:

- 3-column layout
- `貨架位 | 明細SKU | 數量`
- `箱號` shown in row 1
- `SKU總數 | 總數量` shown together in row 2
- page info shown in row 2
- per-box page numbering
- per-box total quantity calculation
- 14-row page capacity
- no empty spacer row between page blocks

---

## Practical Assessment

For current warehouse use, this version is:

- usable
- stable enough for daily operation
- not yet deeply refactored

That means it is suitable for real work, but still has room for code cleanup and structural improvement.

---

## Recommended Next Upgrade

Recommended future improvements:

- refactor duplicated validation logic into one routine
- replace repeated box searching with a dictionary-based structure
- centralize all layout constants
- add a full error handler for unexpected runtime failures
- optionally show all invalid rows at once instead of stopping at the first one

---

## Reference

Updated according to the current canvas VBA script and latest user-confirmed layout.

---

**End of README**
