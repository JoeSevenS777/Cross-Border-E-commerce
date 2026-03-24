# Joe77 Procurement & BigSeller Import Automation (VBA)

[![Status](https://img.shields.io/badge/status-stable-brightgreen)]()
[![Excel Version](https://img.shields.io/badge/Excel-VBA-blue)]()
[![Version](https://img.shields.io/badge/version-2026_update-orange)]()
[![License](https://img.shields.io/badge/license-private-lightgrey)]()

---

## Overview
This VBA macro exports purchase data from four sheets:

- Mengjie (萌睫)
- Caijing (采菁)
- Flortte
- Jiazi (夹子)

It generates two types of files:

- **Joe77 Procurement File**
- **BigSeller Import File**

The script processes rows where the `action` column indicates a purchase requirement.

---

## Core Workflow (Original Logic Preserved)

### 1. Entry Point
```
Joe采购单_Bigseller采购导入
```

- Validates workbook
- Validates active sheet
- Disables Excel performance settings
- Routes logic by sheet name

---

### 2. Sheet Routing

| Sheet | Handler |
|------|--------|
| 萌睫 | Handle_MengJie |
| 采菁 | Handle_CaiJing |
| Flortte | Handle_Flortte |
| 夹子 | Handle_JiaZi |

---

### 3. Row Filtering Logic

```
IsPlaceOrderAction(action)
```

Rule:

> Any value containing **"place order"** (case-insensitive) is valid

Examples:
- Place Order
- Place Order - Network
- Place Order - Default

---

### 4. File Naming

```
<prefix> + yyyymmdd + .xlsx
```

Examples:
- Joe77采购单_萌睫20260324.xlsx
- Bigseller采购导入_采菁20260324.xlsx

---

### 5. File Handling

```
EnsureTodayFile(folder, prefix, todayFile)
```

- If today's file exists → reuse
- If old file exists → rename to today
- Prevents duplication

---

## Output Logic by Sheet

### 1. Mengjie & Caijing (Unified)

**Joe77 Output:**

| Column | Field |
|-------|------|
| A | 商品链接 (Product Link) |
| B | 属性SKU (Attribute SKU) |
| C | 数量 (Quantity) |

**Logic:**
```
Export_Mengjie_Joe77
```

---

### 2. Flortte

**Joe77 Output:**

| Column | Field |
|-------|------|
| A | SKU |
| B | 数量 |

**Special Rule:**
- Remove prefix: `MJ-Flortte-`

---

### 3. Jiazi

Same structure as Flortte

**Prefixes removed:**
- MJ-夹子-
- MJ-JiaZi-

---

### 4. BigSeller Import (All Sheets)

```
Export_ImportFile
```

- Fixed **18-column template**
- Only fills:
  - SKU Name
  - Purchase Qty

---

## File Save Location (Updated)

All outputs are saved to:

```
ThisWorkbook.Path
```

✔ No external dependency
✔ Consistent across all sheets

---

## Automation Behavior

- Clears old data but keeps headers
- Reuses existing files
- Suppresses Excel alerts
- Auto-fit columns
- Prompts user to open files after execution

---

## Requirements

- Workbook must be saved
- Required headers:
  - action
  - SKU / 属性SKU / 商品链接 / 数量
- Only works on 4 specific sheets

---

## Design Principles

- **Consistency** → unified output format
- **Flexibility** → fuzzy action matching
- **Stability** → no hardcoded paths
- **Efficiency** → reuse existing files

---

## License
Private internal automation. Not for redistribution.
