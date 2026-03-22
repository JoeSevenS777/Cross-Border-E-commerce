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
* Handles missing files, header variations, and edge cases robustly

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

* Always creates:

```text
Bigseller调拨导入YYYYMMDD.xlsx
```

* Structure:

  * Sheet name: `SKU`
  * Exact required headers

* Auto-fills:

  * SKU
  * Transfer Qty
  * Warehouse fields

* Automatically:

  * deletes same-day duplicate
  * recycles all old dated files

---

### 5. 🇹🇼 Taiwan Warehouse File Handling

* Works with OneDrive file:

```text
台湾仓库调拨表*.xlsm
```

* If missing:

  1. copy latest from local logs
  2. OR create new file with headers

* Workflow:

  * backup to local logs
  * clear old data
  * write new transfer data
  * rename with timestamp:

```text
台湾仓库调拨表YYYYMMDDHHMMSS.xlsm
```

---

### 6. ♻️ Safe File Management

* Uses **Recycle Bin (not permanent delete)**
* Auto-handles:

  * open file conflicts
  * missing files
  * duplicate filenames

---

### 7. 🛡️ Robust Error Handling

Stops execution with clear messages when:

* folder missing
* sheet missing
* header missing

Ensures:

> No silent failure. No corrupted output.

---

### 8. ⚙️ Self-Healing Design

The system automatically recovers from:

* missing Bigseller files
* missing Taiwan files
* empty datasets

This makes it suitable for **daily operational use without manual prep**.

---

## 🧩 Data Mapping

### Bigseller

| Source | Target       |
| ------ | ------------ |
| SKU    | Merchant SKU |
| 数量     | Transfer Qty |

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
3. Create Bigseller file
4. Recycle old Bigseller files
5. Backup Taiwan file
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

* **Idempotent** → safe to run repeatedly
* **Deterministic** → same input = same output
* **Fail-fast** → stops early on errors
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
  * converts numeric values

* Rows with:

  * `transfer` + empty qty → still exported

---

## ✅ Status

`PRODUCTION READY` `STABLE` `FIELD-TESTED`

---

## 🧭 Future Enhancements (Optional)

* Keep last N Bigseller files instead of deleting all
* Add execution log file
* Add UI button inside Excel ribbon

---

**Author:** Joe (Cross-border E-commerce Automation)
