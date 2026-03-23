# Taiwan Outer Package Label Macro - Active Workbook Version (Updated)

Automated outer-package label printing for exported packing workbooks in Excel.

![EXCEL](https://img.shields.io/badge/EXCEL-AUTOMATION-4d4d4d?style=for-the-badge)
![LANGUAGE](https://img.shields.io/badge/LANGUAGE-VBA-d4af37?style=for-the-badge)
![FEATURE](https://img.shields.io/badge/FEATURE-OUTER_PACKAGE_LABEL-1e88e5?style=for-the-badge)
![WORKFLOW](https://img.shields.io/badge/WORKFLOW-ACTIVE_WORKBOOK-43a047?style=for-the-badge)
![GROUPING](https://img.shields.io/badge/GROUPING-BOX_NUMBER-616161?style=for-the-badge)
![OUTPUT](https://img.shields.io/badge/OUTPUT-PRINT_PREVIEW-ff8f00?style=for-the-badge)
![STATUS](https://img.shields.io/badge/STATUS-STABLE-52c41a?style=for-the-badge)

---

## 🔄 Key Updates (IMPORTANT)

### 1. Dynamic Header Detection

The macro **no longer depends on fixed column positions**.

| Field    | Header Name |
| -------- | ----------- |
| SKU      | SKU编号       |
| Quantity | 实发数量        |
| Box No   | 箱號 / 箱号     |

👉 Columns can be reordered safely as long as headers remain unchanged.

---

### 2. Smart Selection Logic (NEW ⭐)

When selecting cells in the **箱号 column**:

#### Case A — Normal selection

* Only selected cells are processed

#### Case B — Whole column selection

* Treated as: **all non-empty 箱号 cells within data range**
* Empty cells are automatically ignored

👉 This matches real workflow:

> “Select entire column = process all existing boxes”

---

### 3. Strict Data Validation (Enhanced)

Validation is applied to **selected non-empty 箱号 rows only**.

### Validation Rules

| Field | Rule            | Result |
| ----- | --------------- | ------ |
| SKU编号 | Cannot be empty | ❌ Stop |
| 实发数量  | Cannot be empty | ❌ Stop |
| 实发数量  | Must be numeric | ❌ Stop |
| 实发数量  | Must be > 0     | ❌ Stop |
| 实发数量  | Must be integer | ❌ Stop |

### Example Warning

```
实发数量 不是整数（行 25，箱号 03）。请修正后再打印。
```

👉 Behavior:

* Blank rows are ignored
* Only **real data rows are validated**
* First invalid row stops the process immediately

---

## ✨ Features

### 1. Active Workbook Workflow

* Runs on **current active workbook**
* No data copying required

---

### 2. Box-Based Grouping

* Groups rows by **箱号**
* One label per box

---

### 3. Intelligent Filtering

Automatically excludes:

* Header row
* Empty 箱号 cells
* Invalid rows (blocked by validation)

---

### 4. Box Number Normalization

Outputs:

* 01, 02, 03

---

### 5. Multi-Page Support

* Automatically paginates large boxes
* Repeats header per page

---

### 6. Print Preview Workflow

* Always preview before printing

---

## 🧾 Label Content

Each label contains:

* 箱号
* SKU数
* 明细SKU * 数量

Example:

```
箱号：01
SKU数：2
明细SKU * 数量：
MJ-AAA  *  30
MJ-BBB  *  10
```

---

## 🧩 Workflow

1. Open exported packing workbook
2. Select cells in **箱号 column** (or whole column)
3. Run macro
4. Validation runs (only on non-empty rows)
5. Labels generated → PrintPreview

---

## ⚠️ Error Handling Summary

| Scenario               | Behavior                               |
| ---------------------- | -------------------------------------- |
| Missing header         | Stop + show missing headers            |
| Empty selection        | Stop                                   |
| Invalid row (selected) | Stop + show row + box                  |
| Whole column selection | Ignore blanks, validate real data only |
| No valid rows          | Stop                                   |

---

## 🧠 Design Philosophy

This macro enforces **real-world warehouse constraints**:

* No fractional shipments
* No zero shipment
* No dirty data allowed into labels

At the same time, it improves usability:

* Whole-column selection behaves intuitively
* Blank rows are ignored
* Only meaningful data is validated

👉 Goal: balance **data integrity + operational efficiency**

---

## 📌 Status

**Stable (Smart Selection + Strict Validation Version)**

---

## 📎 Original Reference

Based on previous version: fileciteturn1file0

---

## 🚀 Recommended Next Upgrade

* Show ALL invalid rows at once
* Auto-highlight invalid rows
* Export error report

---

**End of README**
