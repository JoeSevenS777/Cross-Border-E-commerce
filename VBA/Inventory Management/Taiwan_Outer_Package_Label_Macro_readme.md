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

## 🔄 Key Update (IMPORTANT)

### Dynamic Header Detection

The macro **no longer depends on fixed column positions**.

Instead, it dynamically detects columns based on header names:

| Field    | Header Name |
| -------- | ----------- |
| SKU      | SKU编号       |
| Quantity | 实发数量        |
| Box No   | 箱號 / 箱号     |

👉 This means:

* You can reorder columns freely
* The macro will still work
* As long as header names remain unchanged

---

## 🧪 Data Validation (NEW)

Before generating labels, the macro now performs **strict validation on 实发数量**.

### Validation Rules

| Rule            | Condition | Result           |
| --------------- | --------- | ---------------- |
| Must be numeric | 非数字       | ❌ Stop + warning |
| Must be > 0     | ≤ 0       | ❌ Stop + warning |
| Must be integer | 有小数       | ❌ Stop + warning |

### Example Warning

```
实发数量 不是整数（行 25，箱号 03）。请修正后再打印。
```

👉 Behavior:

* Validation only checks **selected box rows**
* Stops immediately when error is found
* Prevents incorrect label generation

---

## ✨ Features

### 1. Active Workbook Workflow

* Macro runs on **current active workbook**
* No need to copy data into master file

---

### 2. Box-Based Grouping

* Groups all rows by **箱号**
* One label per box

---

### 3. Intelligent Filtering

Automatically excludes:

* Header row
* Empty box numbers
* Invalid rows (missing SKU / quantity / zero / decimal)

---

### 4. Box Number Normalization

Outputs:

* 01, 02, 03

---

### 5. Multi-Page Support

* Automatically paginates large boxes
* Repeats header on each page

---

### 6. Print Preview Workflow

* Always opens preview first
* Prevents printing errors

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
2. Select cells in **箱号 column**
3. Run macro
4. Validation runs
5. Labels generated → PrintPreview

---

## ⚠️ Error Handling Summary

| Scenario           | Behavior                         |
| ------------------ | -------------------------------- |
| Missing header     | Stop + show which header missing |
| No valid selection | Stop                             |
| Quantity invalid   | Stop + show row + box            |
| No valid rows      | Detailed diagnostic per box      |

---

## 🧠 Design Philosophy

This macro enforces **real-world warehouse constraints**:

* No fractional shipments
* No zero shipment
* No dirty data allowed into labels

👉 Goal: eliminate downstream errors in logistics & receiving

---

## 📌 Status

**Stable (Enhanced Validation Version)**

---

## 📎 Original Reference

Based on previous version: fileciteturn1file0

---

## 🚀 Recommended Next Upgrade

* Show ALL invalid rows at once (instead of first)
* Auto-highlight invalid rows
* Export error report

---

**End of README**
