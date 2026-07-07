# UpdateInventoryNEW – Automated Inventory Synchronization

Reliable, header-aware, language-flexible stock updating for Excel

![Excel](https://img.shields.io/badge/Excel-Automation-217346?style=for-the-badge\&logo=microsoft-excel\&logoColor=white)
![VBA](https://img.shields.io/badge/Language-VBA-yellow?style=for-the-badge)
![Inventory](https://img.shields.io/badge/Feature-Inventory%20Sync-blue?style=for-the-badge)
![SKU](https://img.shields.io/badge/SKU%20%2B%20Warehouse-Matching-orange?style=for-the-badge)
![Data](https://img.shields.io/badge/Source-English%20%2F%20Chinese-success?style=for-the-badge)
![Status](https://img.shields.io/badge/Status-Stable-brightgreen?style=for-the-badge)

---

An advanced VBA macro that automatically updates your **Inventory List** by synchronizing data from the latest stock workbook (English or Chinese).

This script supports:

* Header detection in **both English & Chinese**
* **SKU + Warehouse** matching
* Shelf synchronization
* Auto-cleaning old values
* Protection-aware worksheet handling
* Fast performance via dictionary lookups
* **Available Stock offset (−10,000) for non-Taiwan warehouses only**
* Multi-sheet updating
* One-click UI button

---

## ✨ What This Script Does

### 1. Locates the newest inventory data file

Searches the configured folder for the most recently modified file:

* `Inventory_List*.xlsx`
* `库存清单*.xlsx`

The file is opened safely in **read-only mode**.

---

### 2. Detects headers automatically (fuzzy matching)

Understands both English and Chinese header names in the source workbook:

| Field | English Header | Chinese Header |
| --- | --- | --- |
| SKU | SKU Name | SKU编号 |
| Warehouse | Warehouse | 仓库 |
| Shelf | Shelf | 货架位 |
| Available | Available for whole warehouse / Available Stock | 整仓可用 |
| On the Way | On the Way | 在途中 |
| Order Allocated | Order Allocated | 订单已锁 |
| Daily Sales | Forecasted Daily Sales | 预测日销量 |

The script normalizes headers and ignores spaces, punctuation, line breaks, and capitalization.

---

### 3. Loads the source data into dictionaries by `SKU + Warehouse`

Data is cached using a composite key:

```text
SKU||Warehouse
```

This allows the same SKU to exist in multiple warehouses without overwriting each other.

The script stores:

* Shelf
* Available Stock
* On the Way
* Order Allocated
* Forecasted Daily Sales

This guarantees fast lookup even with large datasets.

---

### 4. Updates multiple inventory sheets

Each run automatically updates the following sheets if they exist:

* `采菁`
* `萌睫`
* `Flortte`
* `MJSP`
* `魔仙`
* `夹子`

No active-sheet dependency exists.

---

### 5. Clears and refills inventory columns

Before inserting new data, the script clears these target columns:

* Shelf
* Available Stock
* On the Way
* Order Allocated
* Daily Sales

Then it refills them by matching each row with:

* `SKU`
* `Warehouse`

Matching is case-insensitive because the dictionaries use `vbTextCompare`.

---

### 6. Applies the “Available Stock − 10,000” rule with Taiwan exemption

After the new values are written, the script loops through `Available Stock` and applies this rule:

* If the row’s `Warehouse` is `Taiwan` (case-insensitive), **do not subtract 10,000**
* For all other warehouses, subtract **10,000** from numeric `Available Stock` values

Examples:

* `Taiwan` → no subtraction
* `taiwan` → no subtraction
* `TAIWAN` → no subtraction
* `Default Warehouse` → subtract 10,000

This keeps the Taiwan warehouse exempt while preserving the stock buffer rule for other warehouses.

---

### 7. Recalculates formulas and restores protection

After updating each sheet:

* Worksheet formulas are recalculated
* Sheet protection is restored
* Application states are returned to normal

The script runs silently. Messages appear only when an error occurs.

---

## 🛠 Setup

### Folder setting

Update this constant in the VBA module:

```vb
Private Const DATA_FOLDER As String = "C:\Users\zouzh\Downloads"
```

### Source files expected

Any `.xlsx` file whose name starts with:

* `Inventory_List`
* `库存清单`

The most recently modified file is chosen automatically.

### Required source headers

The source workbook’s first sheet should contain these fields:

* SKU Name / SKU编号
* Warehouse / 仓库
* Shelf / 货架位
* Available for whole warehouse / 整仓可用 / Available Stock
* On the Way / 在途中
* Order Allocated / 订单已锁
* Forecasted Daily Sales / 预测日销量

### Required target headers

Each target sheet should contain these headers:

* SKU
* Warehouse
* Shelf
* Available Stock
* On the Way
* Order Allocated
* Daily Sales

---

## 🚀 How to Use

1. Open **Inventory List.xlsm**
2. Click the **Update** button next to the SKU header, or run `UpdateInventoryNEW` manually
3. The configured sheets are updated automatically
4. No confirmation popup appears unless something goes wrong

---

## ✔ Supported Scenarios

* English source workbook
* Chinese source workbook
* Mixed-language environments
* Protected sheets
* Missing or reordered columns
* Same SKU appearing in multiple warehouses
* Case differences in warehouse names such as `Taiwan`, `taiwan`, `TAIWAN`
* Large datasets

---

## 📄 License

MIT License — suitable for business and operational automation.
