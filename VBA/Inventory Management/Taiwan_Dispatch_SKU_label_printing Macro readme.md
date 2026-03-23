# Taiwan Label Print Macro

Automated per-SKU label printing for Taiwan warehouse operations in Excel.

![EXCEL](https://img.shields.io/badge/EXCEL-AUTOMATION-4d4d4d?style=for-the-badge)
![LANGUAGE](https://img.shields.io/badge/LANGUAGE-VBA-d4af37?style=for-the-badge)
![FEATURE](https://img.shields.io/badge/FEATURE-SKU_LABEL_PRINT-1e88e5?style=for-the-badge)
![WORKFLOW](https://img.shields.io/badge/WORKFLOW-WAREHOUSE_DISPATCH-43a047?style=for-the-badge)
![SELECTION](https://img.shields.io/badge/SELECTION-ROW_%2F_SKU_COLUMN-616161?style=for-the-badge)
![OUTPUT](https://img.shields.io/badge/OUTPUT-PRINT_PREVIEW-ff8f00?style=for-the-badge)
![STATUS](https://img.shields.io/badge/STATUS-STABLE-52c41a?style=for-the-badge)

A VBA macro that prints **one SKU label per page** from the Taiwan dispatch sheet.

---

## ✨ Key Updates (Important)

* ✅ Uses **Sheet1** as the source sheet
* ✅ Uses **PrintPreview** as the output sheet
* ✅ Reads columns by **header name (not fixed column letters)**
* ✅ More robust to future layout changes

---

## 🧾 Required Headers (Sheet1)

The macro dynamically finds these headers in **Row 1**:

| Header | Description             |
| ------ | ----------------------- |
| SKU编号  | SKU identifier          |
| GTIN   | Barcode number          |
| 台湾货架位  | Shelf location          |
| 调拨数量   | Quantity used for label |

> ⚠️ Header names must match exactly.

---

## 🏷 Label Content

Each label contains:

* 货架位
* SKU
* GTIN
* 数量（来自“调拨数量”）

Example:

```
货架位：C-03-05
SKU：MJ-萌睫尚品-自粘云朵
GTIN：6942265706768
数量：10
```

---

## 🔁 Workflow

1. Open the workbook
2. Go to **Sheet1**
3. Select:

   * rows OR
   * SKU cells OR
   * full SKU column
4. Run macro: `PrintSelectedLabelsPreview`
5. Labels are generated in `PrintPreview`
6. Excel opens **Print Preview**

---

## ⚙️ Behavior Rules

The macro will:

* ignore header row
* ignore empty SKU rows
* avoid duplicate row processing
* preserve GTIN as full text (no scientific notation)
* generate **1 label = 1 page**

---

## 🖨 Layout

Designed for:

```
100mm × 100mm label paper
```

* clear borders
* large readable text
* warehouse-friendly layout

---

## 🧠 Design Logic

This macro is optimized for warehouse operations where:

* each SKU needs its own label
* shelf location must be visible
* operators need fast identification
* printing must be previewed before execution

---

## 🚀 Recommended Use Cases

* Taiwan warehouse transfer
* SKU-level dispatch labeling
* product packaging labels
* receiving & shelving support

---

## 📌 Notes

* Must run macro on **Sheet1**
* `PrintPreview` sheet must exist
* Headers must not be renamed

---

## Status

**Stable (Header-driven version)**

* resilient to column movement
* selection-aware
* production-ready for daily warehouse use
