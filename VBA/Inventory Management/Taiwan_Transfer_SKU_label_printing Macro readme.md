# Taiwan Label Print Macro

Automated per-SKU label printing for Taiwan warehouse operations in Excel.

![EXCEL](https://img.shields.io/badge/EXCEL-AUTOMATION-4d4d4d?style=for-the-badge)
![LANGUAGE](https://img.shields.io/badge/LANGUAGE-VBA-d4af37?style=for-the-badge)
![FEATURE](https://img.shields.io/badge/FEATURE-SKU_LABEL_PRINT-1e88e5?style=for-the-badge)
![WORKFLOW](https://img.shields.io/badge/WORKFLOW-WAREHOUSE_DISPATCH-43a047?style=for-the-badge)
![SELECTION](https://img.shields.io/badge/SELECTION-ROW_%2F_SKU_COLUMN-616161?style=for-the-badge)
![OUTPUT](https://img.shields.io/badge/OUTPUT-PRINT_PREVIEW-ff8f00?style=for-the-badge)
![STATUS](https://img.shields.io/badge/STATUS-STABLE-52c41a?style=for-the-badge)

A VBA macro that prints **one SKU label per page** and records printing status to prevent duplicate shipment risk.

---

## ✨ Key Features

### 1. One SKU = One Label

Each selected SKU generates **one label per page**.

---

### 2. Header-Driven (Robust)

The macro detects columns by **header names**, not fixed positions.

Required headers:

| Header |
| ------ |
| SKU编号  |
| GTIN   |
| 台湾货架位  |
| 调拨数量   |
| 标签打印状态 |

---

### 3. Duplicate Print Prevention (Core Feature)

Before printing, the macro checks:

* if any selected rows are already marked as printed

If found, it shows:

```
行 3,4,5 已打印，请确认是否继续打印？
```

* Yes → continue
* No → stop

---

### 4. Print Confirmation Control

After closing Print Preview, the macro asks:

```
请确认：以上标签是否已经成功打印？
点击“是”将标记为“已打印”。
```

If **Yes**:

* writes timestamp into `标签打印状态`

Format:

```
mm/dd_hh:mm已打印
```

Example:

```
03/06_12:11已打印
```

If **No**:

* no changes are made

---

### 5. Conditional Formatting Recommended

Color highlighting is NOT handled by VBA.

Instead, use Excel **Conditional Formatting**:

Example formula:

```
=ISNUMBER(SEARCH("已打印", $J2))
```

(Replace `J` with your 标签打印状态 column)

Benefits:

* automatically reversible
* no hard formatting
* consistent visual control

---

### 6. Selection-Aware Workflow

Supports:

* selecting rows
* selecting SKU cells
* selecting ranges

Automatically:

* ignores header row
* ignores empty SKU
* avoids duplicate row processing

---

### 7. Print Preview First

Workflow:

1. generate labels
2. open Print Preview
3. user prints or checks
4. confirmation step updates status

---

## 🏷 Label Content

Each label contains:

* 货架位
* SKU
* GTIN
* 数量

---

## ⚙️ Workflow Summary

1. Go to **Sheet1**
2. Select SKU rows/cells
3. Run macro `PrintSelectedLabelsPreview`
4. Preview opens
5. Print if needed
6. Confirm → mark as printed

---

## 🧠 Control Design

This macro adds **operational control**:

* prevents accidental duplicate printing
* adds traceability via timestamp
* enforces human confirmation step

---

## 🚀 Use Cases

* Taiwan warehouse transfer
* SKU labeling
* dispatch preparation
* receiving & shelving

---

## Status

**Stable (Control-enabled version)**

* header-driven
* duplicate warning
* print confirmation
* timestamp tracking
