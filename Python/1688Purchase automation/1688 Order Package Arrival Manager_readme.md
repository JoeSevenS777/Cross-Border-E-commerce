# 1688 Package Arrival Manager

**Order ⇄ Tracking Sync & Arrival Updates**

![Python](https://img.shields.io/badge/Python-3.8%2B-blue)
![OpenPyXL](https://img.shields.io/badge/openpyxl-Excel%20I%2FO-orange)
![Workflow](https://img.shields.io/badge/Workflow-1688%20Logistics-informational)
![Status](https://img.shields.io/badge/Status-Production-brightgreen)

Automates **on-the-way order & package tracking maintenance** for 1688 exports.

**Key features:**

* Converts noisy 1688 exports into a **clean ledger**
  → one row per **(订单编号, 运单号)**
* Preserves manual fields: **到达情况 / 备注**
* Supports **paste-to-update arrival confirmations** (WeChat / text blocks)
* Highlights:

  * duplicate **订单编号** (one order, multiple packages)
  * duplicate **运单号** (one package, multiple orders)

---

## Overview

This script processes Excel files in the **same folder as the script** and maintains a clean tracking workbook.

* **Input**: latest 1688 export workbook (filename is variable)
* **Output**: `Order_Package_arrival_Confirmation.xlsx`
* **Archive**: parsed exports are moved to `./parsed files/`

The clean workbook stays **human-friendly, auditable, and safe** for daily operations.

---

## What This Script Does

### Mode 1 — Sync from latest 1688 export

* Finds the **most recently modified** `.xlsx` in the base folder
  (excluding `Order_Package_arrival_Confirmation.xlsx` and `~$*`)
* Counts **distinct 订单编号** and asks for confirmation before syncing
* **Normalizes 运单号 values** to avoid Excel numeric issues (e.g. `435000…` vs `435000….0`) before building keys
* Extracts only required columns:

  * 订单编号
  * 卖家公司名
  * 订单状态
  * 订单创建时间
  * 订单付款时间
  * 物流公司
  * 运单号
* Builds a clean dataset where:

> **each row = one (订单编号, 运单号)**

* Preserves existing:

  * 到达情况
  * 备注
* Applies formatting:

  * Duplicate **订单编号** → highlighted in red
  * Duplicate **运单号** → highlighted in red
  * Header totals:

    * `订单编号（N）` = unique order count
    * `运单号（M）` = unique tracking count
  * Auto-adjusted column widths
* Moves the parsed export into `./parsed files/`
* **On successful completion, automatically opens** `Order_Package_arrival_Confirmation.xlsx`

---

### Mode 2 — Paste arrival info

* Waits for pasted text blocks
* Updates the ledger by **运单号**
* Supports text blocks like:

```text
【承运公司】圆通速递(YTO)
【运单号】YT7598163190233
恒志达齐
```

* Updates **all matching rows**
  (important when one 运单号 serves multiple orders)
* **Normalizes pasted 运单号** to ensure matching even if Excel stored it as a number
* Sets `到达情况` using conservative inference:

  * contains `退款` → 退款
  * contains `退货` → 退货
  * contains `异常` or `丢` → 异常
  * otherwise → 已到达
* Appends new content into `备注` (never overwrites existing notes)
* Prints any **unmatched 运单号** to ensure nothing is silently missed
* **On successful completion, automatically opens** `Order_Package_arrival_Confirmation.xlsx`

---

## Folder Structure

```text
Package_Arrival_Confirmation/
│
├─ Package_Arrival_Management.py
├─ Package_Arrival_Management.bat
├─ Order_Package_arrival_Confirmation.xlsx
├─ <latest_1688_export>.xlsx
└─ parsed files/
```

---

## Requirements

* Python **3.8+**
* `openpyxl`

Install dependency:

```bash
pip install openpyxl
```

---

## Usage

### Recommended (Windows)

Double-click:

```text
Package_Arrival_Management.bat
```

---

### Mode 1 — Sync

1. Export **on-the-way orders** from 1688 (`.xlsx`)
2. Place the export file in the same folder as the script
3. Run the tool → choose `1`
4. Confirm when prompted:

```text
待收货订单总有 41，
确认请按Y/y同步，否则按其他任意键退出:
```

Only `Y` / `y` will proceed.

---

### Mode 2 — Arrival update

1. Run the tool → choose `2`
2. Paste arrival text blocks
3. Press **Enter twice**
4. The tool updates **到达情况 / 备注** and reports unmatched 运单号

---

## Important Notes

### 1) Excel must be **CLOSED** (or auto-closed)

If the clean workbook is open, Windows

If `Order_Package_arrival_Confirmation.xlsx` is open:

* Windows locks the file
* The script exits safely
* No partial writes occur

**Solution:** close Excel and run again.

---

### 2) No missing IDs guarantee

**Design rule:**

> Do **not** miss any 订单编号 or 运单号.

* One order → multiple packages
  → duplicate **订单编号** rows
* One package → multiple orders
  → duplicate **运单号** rows

Both are intentional and highlighted.

---

## Output Rules (Data Model)

Each output row represents:

```text
(订单编号, 运单号)
```

This is the **only reliable structure** that supports:

* multi-package orders
* shared packages across orders
* zero loss of tracking numbers

---

## Troubleshooting

### “No new export workbook found”

Possible reasons:

* You already moved the export into `./parsed files/`
* No new export exists in the base folder

**Solution:**

1. Export again from 1688
2. Place the `.xlsx` next to the script
3. Run **Mode 1**

---

### “Permission denied”

Cause:

* Clean workbook is open in Excel

**Solution:**

* Close Excel
* Run again

---

## License

Internal workflow tool (adjust as needed).
