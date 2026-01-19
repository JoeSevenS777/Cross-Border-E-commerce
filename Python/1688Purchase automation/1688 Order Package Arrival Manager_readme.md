# 1688 Package Arrival Manager (Order ⇄ Tracking Sync & Arrival Updates)

![Python](https://img.shields.io/badge/Python-3.8%2B-blue)
![OpenPyXL](https://img.shields.io/badge/openpyxl-Excel%20I%2FO-orange)
![Workflow](https://img.shields.io/badge/Workflow-1688%20Logistics-informational)
![Status](https://img.shields.io/badge/Status-Production-brightgreen)

Automates **on-the-way order & package tracking maintenance** for 1688 exports.

- Converts the noisy 1688 export into a **clean ledger**: one row per **(订单编号, 运单号)** relationship  
- Preserves your manual fields (**到达情况 / 备注**) across updates  
- Supports **paste-to-update arrival confirmations** (WeChat / text blocks)  
- Highlights **duplicate 订单编号** (multi-package orders) and **duplicate 运单号** (shared packages across orders)

---

## Overview

This script processes Excel files in the **same folder as the script** and maintains a clean tracking workbook:

- **Input**: latest 1688 export workbook (filename is variable)
- **Output**: `Order_Package_arrival_Confirmation.xlsx`
- **Archive**: parsed exports are moved to `./parsed files/`

The clean workbook stays human-friendly and operationally safe.

---

## What This Script Does

### Mode 1 — Sync from latest 1688 export
- Finds the **most recently modified** `.xlsx` in the base folder (excluding `Order_Package_arrival_Confirmation.xlsx` and `~$*`)
- Counts total **distinct 订单编号** and asks for confirmation before updating  
- Extracts only these columns from the export:
  - 订单编号 / 卖家公司名 / 订单状态 / 订单创建时间 / 订单付款时间 / 物流公司 / 运单号
- Builds a clean dataset where **each row = one (订单编号, 运单号)**
- Preserves existing:
  - 到达情况
  - 备注
- Applies formatting:
  - Duplicate **订单编号** highlighted in red
  - Duplicate **运单号** highlighted in red
  - Header counts:
    - `订单编号（N）` = unique order count
    - `运单号（M）` = unique tracking count
  - Auto-adjust column widths
- Moves the parsed export into `./parsed files/`

### Mode 2 — Paste arrival info
- Waits for pasted text blocks, then updates the ledger by **运单号**
- Supports blocks like:
【承运公司】圆通速递(YTO)
【运单号】YT7598163190233
恒志达齐

- Updates all matching rows (important when **one 运单号 serves multiple orders**)
- Sets `到达情况` with conservative inference:
- contains `退款` → 退款
- contains `退货` → 退货
- contains `异常` or `丢` → 异常
- otherwise → 已到达
- Appends new note content into `备注` (does not overwrite existing info)
- Prints any unmatched 运单号 to ensure nothing is silently missed

---

## Folder Structure

Place everything in a single base folder:

Package_Arrival_Confirmation/
│
├─ Package_Arrival_Management.py
├─ Package_Arrival_Management.bat
├─ Order_Package_arrival_Confirmation.xlsx
├─ <latest_1688_export>.xlsx
└─ parsed files/

---

## Requirements

- Python **3.8+**
- `openpyxl`

Install:

```bash
pip install openpyxl
