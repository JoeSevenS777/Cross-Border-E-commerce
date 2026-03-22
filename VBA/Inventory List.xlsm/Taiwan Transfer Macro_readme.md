# Taiwan Transfer Macro

Automated Transfer Workflow for Cross-border Inventory → Bigseller ERP + Taiwan Warehouse

---

### 🏷️ Badges

![Excel](https://img.shields.io/badge/Excel-Automation-green)
![VBA](https://img.shields.io/badge/Language-VBA-yellow)
![Status](https://img.shields.io/badge/Status-Stable-brightgreen)
![System](https://img.shields.io/badge/System-Inventory%20Transfer-blue)
![Platform](https://img.shields.io/badge/Platform-Bigseller%20ERP-orange)
![Warehouse](https://img.shields.io/badge/Warehouse-Taiwan-purple)
![Feature](https://img.shields.io/badge/Feature-Header%20Matching-lightgrey)
![Feature](https://img.shields.io/badge/Feature-Visible%20Rows%20Only-lightgrey)
![Design](https://img.shields.io/badge/Design-Self--Healing-success)

---

## 📌 Overview

A **production-grade VBA automation system** that:

* Extracts transfer data from multiple inventory sheets
* Generates **Bigseller import files automatically**
* Updates **Taiwan warehouse transfer records in OneDrive**
* Handles header variations and operational edge cases robustly

Designed for **real-world cross-border e-commerce operations** where reliability and repeatability matter.

---

## ✨ Core Features

### 1. 🔄 End-to-End Automation

* One-click execution
* Multi-sheet scanning
* Full pipeline: **detect → extract → generate → archive → update**

---

### 2. 📊 Intelligent Data Detection

* Scans sheets:

  * 采菁 / 萌睫 / Flortte / 魔仙 / 夹子 / 其他备货
* Detects rows where:

  * `action = transfer` (case-insensitive)
* Works only on:

  * **visible rows (respects filters)**

---

### 3. 🧠 Flexible Header Matching

* Does NOT rely on fixed column positions
* Ignores:

  * spaces
  * line breaks
  * casing
* Prevents breakage when columns are rearranged

---

### 4. 📦 Bigseller File Generation (Always Fresh)

* Always creates a uniquely timestamped file:

```text
Bigseller调拨导入YYYYMMDDHHMMSS.xlsx
```

* Structure:

  * Sheet name: `SKU`
  * Exact required headers (row 1):

| Column Order | Header                                                |
| ------------ | ----------------------------------------------------- |
| 1            | *Provisional No.(Required)                            |
| 2            | *Shipping Warehouse(Required)                         |
| 3            | *Receiving Warehouse(Required)                        |
| 4            | *Merchant SKU(Required)                               |
| 5            | *Transfer Qty(Required)                               |
| 6            | Ship Fee                                              |
| 7            | Tracking No.                                          |
| 8            | Cost Allocation Method (Price/Quantity/Volume/Weight) |
| 9            | Estimated Arrival Date                                |
| 10           | Note                                                  |

* Auto-fills per row:

  * `*Provisional No.(Required)` = 1
  * `*Shipping Warehouse(Required)` = default warehouse
  * `*Receiving Warehouse(Required)` = Taiwan
  * `*Merchant SKU(Required)` = SKU
  * `*Transfer Qty(Required)` = 数量 (numeric only; blanks allowed)

* Automatically:

  * creates a fresh file every run
  * avoids same-day conflicts via timestamp
  * recycles all other old `Bigseller调拨导入*.xlsx` files

---

### 5. 🇹🇼 Taiwan Warehouse File Handling

* Works with OneDrive file:

```text
台湾仓库调拨表*.xlsm
```

* Required behavior:

  * the file **must already exist in OneDrive**
  * if it is missing, the macro shows a clear error message and exits

* Workflow:

  * backup the current OneDrive file to local logs
  * clear old data
  * write new transfer data
  * rename with timestamp:

```text
台湾仓库调拨表YYYYMMDDHHMMSS.xlsm
```

---

### 6. ♻️ Safe File Management

* Uses **Recycle Bin** for old Bigseller files
* Auto-handles:

  * open file conflicts
  * duplicate filenames
  * old dated file cleanup

---

### 7. 🛡️ Robust Error Handling

Stops execution with clear messages when:

* folder missing
* sheet missing
* header missing
* required Taiwan OneDrive file missing

Ensures:

> No silent failure. No corrupted output.

---

### 8. ⚙️ Stable Production Design

The final workflow is intentionally split:

* **Bigseller** → always create a fresh timestamped export file
* **Taiwan** → update the existing OneDrive working file, archive a local log copy, then rename the OneDrive file with a timestamp

This makes it suitable for **daily operational use** while keeping the Taiwan side compatible with the OneDrive workflow.

---

## 🧩 Data Mapping

### Bigseller

| Source | Target                  |
| ------ | ----------------------- |
| SKU    | *Merchant SKU(Required) |
| 数量     | *Transfer Qty(Required) |

Additional fixed values:

| Target                         | Value             |
| ------------------------------ | ----------------- |
| *Provisional No.(Required)     | 1                 |
| *Shipping Warehouse(Required)  | default warehouse |
| *Receiving Warehouse(Required) | Taiwan            |

---

### Taiwan Warehouse

| Source      | Target |
| ----------- | ------ |
| SKU         | SKU编号  |
| GTIN        | GTIN   |
| Daily Sales | 预测日销量  |
| Total Stock | 台湾库存   |
| 数量          | 调拨数量   |
| shelf       | 台湾货架位  |

---

## 🚀 Execution Flow

1. Scan all sheets
2. Extract transfer rows
3. Create Bigseller file (timestamped)
4. Recycle old Bigseller files
5. Backup current Taiwan OneDrive file to local logs
6. Update Taiwan file
7. Rename Taiwan file (timestamp)
8. Prompt user to open outputs

---

## 📂 Folder Structure

```text
Inventory Folder
│
├── Inventory List.xlsm
├── Bigseller调拨导入*.xlsx
├── Taiwan transfer logs/
│
└── (OneDrive)
    └── 台湾仓库调拨表*.xlsm
```

---

## 🧪 Design Principles

* **Idempotent enough for daily use**
* **Deterministic** → same input = same output structure
* **Fail-fast** → stops early on missing requirements
* **Minimal dependencies** → no external libraries

---

## 🧠 When to Use This

This macro is ideal if you:

* run daily transfer operations
* use Bigseller ERP
* manage Taiwan warehouse stock
* want to eliminate manual Excel work

---

## 📎 Notes

* Quantity handling:

  * keeps blanks if invalid
  * converts numeric values for Bigseller

* Rows with:

  * `transfer` + empty qty → still exported

* Taiwan prerequisite:

  * `台湾仓库调拨表*.xlsm` must already exist in `C:\Users\zouzh\OneDrive\`
  * the macro will not rebuild it from logs or create a blank replacement

---

## ✅ Status

`PRODUCTION READY` `STABLE` `FIELD-TESTED`

---

## 🧭 Future Enhancements (Optional)

* Add execution log file
* Add UI button inside Excel ribbon
* Add bilingual README (EN + 中文)

---

**Author:** Joe (Cross-border E-commerce Automation)
