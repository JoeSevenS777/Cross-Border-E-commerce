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
* Updates **Taiwan warehouse transfer records in the local inventory folder**
* Archives the previous Taiwan transfer file into a local log folder
* Handles header variations and operational edge cases robustly

Designed for **real-world cross-border e-commerce operations** where reliability and repeatability matter.

---

## 📂 Current Folder Configuration

The macro now uses a local inventory folder:

```text
D:\JoeProgramFiles\inventory
```

The Taiwan log folder is:

```text
D:\JoeProgramFiles\inventory\Taiwan transfer logs\
```

Corresponding VBA constants:

```vba
Private Const INVENTORY_FOLDER As String = "D:\JoeProgramFiles\inventory"
Private Const TAIWAN_LOG_FOLDER As String = "D:\JoeProgramFiles\inventory\Taiwan transfer logs\"
```

Required folders:

```text
D:\JoeProgramFiles\inventory\
D:\JoeProgramFiles\inventory\Taiwan transfer logs\
```

Required Taiwan working file pattern:

```text
台湾仓库调拨表*.xlsm
```

The Taiwan working file must already exist in:

```text
D:\JoeProgramFiles\inventory\
```

Example:

```text
D:\JoeProgramFiles\inventory\台湾仓库调拨表20260512153000.xlsm
```

---

## ✨ Core Features

### 1. 🔄 End-to-End Automation

* One-click execution
* Multi-sheet scanning
* Full pipeline: **detect → extract → generate → archive → update → rename**

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
  * tabs
  * casing
  * non-breaking spaces

* Prevents breakage when columns are rearranged

---

### 4. 📦 Bigseller File Generation (Always Fresh)

* Always creates a uniquely timestamped file:

```text
Bigseller调拨导入YYYYMMDDHHMMSS.xlsx
```

* Structure:

  * Sheet name: `SKU`
  * Exact required headers in row 1:

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

* Works with this transfer file pattern:

```text
台湾仓库调拨表*.xlsm
```

* Required behavior:

  * the file **must already exist** in `D:\JoeProgramFiles\inventory\`
  * if it is missing, the macro shows a clear error message and exits

* Workflow:

  * find the latest valid timestamped Taiwan transfer file
  * back up the current file to local logs
  * clear old data
  * write new transfer data
  * save the updated file
  * rename it with a new timestamp:

```text
台湾仓库调拨表YYYYMMDDHHMMSS.xlsm
```

---

### 6. ♻️ Safe File Management

* Uses **Recycle Bin** for old Bigseller files
* Creates a backup of the Taiwan file before updating it
* Auto-handles:

  * duplicate filenames
  * old dated Bigseller file cleanup
  * open workbook conflicts inside the current Excel instance

---

### 7. 🛡️ Robust Error Handling

Stops execution with clear messages when:

* folder missing
* sheet missing
* header missing
* required Taiwan inventory file missing
* output file cannot be renamed
* old file cannot be moved to Recycle Bin

Ensures:

> No silent failure. No corrupted output.

---

### 8. ⚙️ Stable Production Design

The final workflow is intentionally split:

* **Bigseller** → always create a fresh timestamped export file
* **Taiwan** → update the existing local inventory working file, archive a local log copy, then rename the updated file with a timestamp

This makes it suitable for **daily operational use** while keeping the Taiwan side traceable and recoverable through local logs.

---

## 🧩 Data Mapping

### Bigseller

| Source | Target                  |
| ------ | ----------------------- |
| SKU    | *Merchant SKU(Required) |
| 数量   | *Transfer Qty(Required) |

Additional fixed values:

| Target                         | Value             |
| ------------------------------ | ----------------- |
| *Provisional No.(Required)     | 1                 |
| *Shipping Warehouse(Required)  | default warehouse |
| *Receiving Warehouse(Required) | Taiwan            |

---

### Taiwan Warehouse

| Source      | Target   |
| ----------- | -------- |
| SKU         | SKU编号  |
| GTIN        | GTIN     |
| Daily Sales | 预测日销量 |
| Total Stock | 台湾库存 |
| 数量        | 调拨数量 |
| shelf       | 台湾货架位 |

---

## 🚀 Execution Flow

1. Validate required folders
2. Scan all source sheets
3. Extract visible transfer rows
4. Create Bigseller file (timestamped)
5. Recycle old Bigseller files
6. Find the latest Taiwan inventory file
7. Back up the current Taiwan file to the local log folder
8. Update the Taiwan file with new transfer data
9. Rename the updated Taiwan file with a timestamp
10. Prompt user to open outputs

---

## 📂 Folder Structure

```text
D:\JoeProgramFiles\inventory\
│
├── 台湾仓库调拨表YYYYMMDDHHMMSS.xlsm
│
└── Taiwan transfer logs\
    └── 台湾仓库调拨表YYYYMMDDHHMMSS.xlsm
```

The Bigseller output file is created in the same folder as `Inventory List.xlsm`:

```text
[Folder containing Inventory List.xlsm]\
│
├── Inventory List.xlsm
└── Bigseller调拨导入YYYYMMDDHHMMSS.xlsx
```

---

## 🧪 Design Principles

* **Idempotent enough for daily use**
* **Deterministic** → same input = same output structure
* **Fail-fast** → stops early on missing requirements
* **Traceable** → Taiwan files are backed up before being updated
* **Minimal dependencies** → no external libraries required

---

## 🧠 When to Use This

This macro is ideal if you:

* run daily transfer operations
* use Bigseller ERP
* manage Taiwan warehouse stock
* want to eliminate manual Excel work
* want a local-file workflow with timestamped backups

---

## 📎 Notes

* Quantity handling:

  * keeps blanks if invalid
  * converts numeric values for Bigseller

* Rows with:

  * `transfer` + empty qty → still exported

* Taiwan prerequisite:

  * `台湾仓库调拨表*.xlsm` must already exist in `D:\JoeProgramFiles\inventory\`
  * the macro will not rebuild it from logs
  * the macro will not create a blank Taiwan replacement automatically

* The Taiwan log folder must already exist:

```text
D:\JoeProgramFiles\inventory\Taiwan transfer logs\
```

---

## ⚠️ Known Limitations

* Files opened in another Excel process may still block copy/rename operations
* If the Taiwan file updates successfully but rename fails, the data is already updated but the filename may remain unchanged
* The macro selects the latest Taiwan file based on the numeric timestamp in the filename
* The Taiwan log folder is validated but not automatically created

---

## ✅ Status

`PRODUCTION READY` `STABLE` `FIELD-TESTED`

---

## 🧭 Future Enhancements (Optional)

* Add execution log file
* Automatically create the Taiwan log folder if missing
* Add UI button inside Excel ribbon
* Add bilingual README (EN + 中文)
* Add stricter validation for blank transfer quantities

---

**Author:** Joe (Cross-border E-commerce Automation)
